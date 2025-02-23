using OpenCvSharp;
using SharpDX.Direct2D1;
using SharpDX.Direct3D;
using SharpDX.Direct3D11;
using SharpDX.DXGI;
using SharpDX;
using System.Diagnostics;
using SixLabors.ImageSharp.PixelFormats;
using Compunet.YoloSharp;

using DXDevice = SharpDX.Direct3D11.Device;
using D2D = SharpDX.Direct2D1;
using DWrite = SharpDX.DirectWrite;
using RectF = SharpDX.Mathematics.Interop.RawRectangleF;
using Color4 = SharpDX.Mathematics.Interop.RawColor4;
using CvSize = OpenCvSharp.Size;
using CvPoint = OpenCvSharp.Point;
using SxImage = SixLabors.ImageSharp;
using System.Runtime.CompilerServices;
using System.Buffers;
using Compunet.YoloSharp.Data;
using System.Collections.ObjectModel;

namespace RealObjectDetection;

public partial class Form1 : Form
{
	private bool isDarkMode = true;
	private static YoloPredictor? Yolo;

	public Form1()
	{
		InitializeComponent();
		isDarkMode = !DarkMode.IsLightTheme();
		DarkMode.UseImmersiveDarkMode(Handle, isDarkMode);

		Yolo = new(Properties.Resources.yolo11n);
	}

	private void Form1_Load(object sender, EventArgs e)
	{
#if !USE_HWND_RENDER_TARGET
		// SwapChain description
		var desc = new SwapChainDescription()
		{
			BufferCount = 1,
			ModeDescription = new ModeDescription(ClientSize.Width, ClientSize.Height,
					new SharpDX.DXGI.Rational(60, 1), Format.B8G8R8A8_UNorm),
			IsWindowed = true,
			OutputHandle = Handle,
			SampleDescription = new SampleDescription(1, 0),
			SwapEffect = SwapEffect.Discard,
			Usage = Usage.RenderTargetOutput
		};

		DXDevice.CreateWithSwapChain(DriverType.Hardware, DeviceCreationFlags.BgraSupport, desc, out device, out swapChain);
		var context = device.ImmediateContext;

		// Ignore all windows events
		var factory = swapChain.GetParent<SharpDX.DXGI.Factory>();
		factory.MakeWindowAssociation(Handle, WindowAssociationFlags.IgnoreAll);

		// New RenderTargetView from the backbuffer
		backBuffer = Texture2D.FromSwapChain<Texture2D>(swapChain, 0);
		renderView = new RenderTargetView(device, backBuffer);

		// Create Direct2D factory
		d2dFactory = new D2D.Factory();
		surface = backBuffer.QueryInterface<Surface>();
		renderTarget = new RenderTarget(d2dFactory, surface,
			new RenderTargetProperties(new PixelFormat(Format.Unknown, D2D.AlphaMode.Premultiplied)));
#else
		var hwnd_prop = new HwndRenderTargetProperties
		{
			Hwnd = Handle,
			PixelSize = new Size2(ClientSize.Width, ClientSize.Height)
		};
		d2dFactory = new D2D.Factory();
		var render_prop = new RenderTargetProperties
		{
			DpiX = 96,
			DpiY = 96,
		};
		renderTarget = new D2D.WindowRenderTarget(d2dFactory, render_prop, hwnd_prop);
#endif
		whiteBrush = new SolidColorBrush(renderTarget, new Color4(1, 1, 1, 1));
		greenBrush = new SolidColorBrush(renderTarget, new Color4(0, 0.75f, 0, 1));

		dwriteFactory = new DWrite.Factory();
		fontArial20 = new(dwriteFactory, "Arial", 20);
		fontArial30 = new(dwriteFactory, "Arial", 30);

		frame = new Mat();
	}

	private void Form1_FormClosed(object sender, FormClosedEventArgs e)
	{
		isVideoStopped = true;
		renderTarget?.Dispose();
		dwriteFactory?.Dispose();
		d2dFactory?.Dispose();
#if !USE_HWND_RENDER_TARGET
		swapChain?.Dispose();
		device?.Dispose();
#endif
	}

	private void Form1_Shown(object sender, EventArgs e)
	{
		var openFileDlg = new OpenFileDialog
		{
			Title = "Select a file",
			Filter = "Media Files(*.WMV;*.MP4;*.AVI)|*.WMV;*.MP4;*.AVI"
		};
		var result = openFileDlg.ShowDialog();
		if (result == DialogResult.Cancel) {
			Close();
			return;
		}

		videoCapture = new(openFileDlg.FileName);
		if (!videoCapture.IsOpened()) {
			MessageBox.Show("Unexpected error: Unable to play this file", "Error",
				MessageBoxButtons.OK, MessageBoxIcon.Error);
			Close();
			return;
		}

		var videoSize = new System.Drawing.Size(videoCapture.FrameWidth, videoCapture.FrameHeight);
		if (videoSize.Width >= 2560 || videoSize.Height >= 1440)
			ClientSize = new System.Drawing.Size(videoSize.Width / 2, videoSize.Height / 2);
		else if (videoSize.Width > 1280 || videoSize.Height > 720)
			ClientSize = new System.Drawing.Size((int)(videoSize.Width / 1.5f), (int)(videoSize.Height / 1.5f));
		else
			ClientSize = new System.Drawing.Size(videoSize.Width, videoSize.Height);
		CenterToScreen();

		isVideoStopped = false;
	}

	private void Form1_SizeChanged(object sender, EventArgs e)
	{
#if !USE_HWND_RENDER_TARGET
		if (swapChain == null)
			return;
		device?.ImmediateContext.ClearState();
		renderTarget?.Dispose();
		backBuffer?.Dispose();
		renderView?.Dispose();
		surface?.Dispose();
		swapChain?.ResizeBuffers(1, ClientSize.Width, ClientSize.Height, Format.B8G8R8A8_UNorm, SwapChainFlags.AllowModeSwitch);
		backBuffer = Texture2D.FromSwapChain<Texture2D>(swapChain, 0);
		renderView = new RenderTargetView(device, backBuffer);
		surface = backBuffer.QueryInterface<Surface>();
		renderTarget = new RenderTarget(d2dFactory, surface,
			new RenderTargetProperties(new PixelFormat(Format.Unknown, D2D.AlphaMode.Premultiplied)));
#else
		renderTarget?.Resize(new Size2(ClientSize.Width, ClientSize.Height));
#endif
	}

	private void Form1_KeyPress(object sender, KeyPressEventArgs e)
	{

	}

	private long frame_time = 0;
	private String str_fps = "";
	volatile static bool isDetecting = false;
	private static readonly object lockBox = new();

	public void Render(long ms)
	{
		if (isVideoStopped)
			return;

		//var context = device!.ImmediateContext;
		//context.ClearRenderTargetView(renderView, new Color4(0, 0, 1, 1));

		var elapsed = ms - frame_time;
		double fps = 1000.0 / elapsed;
		if (fps < videoCapture!.Fps * 1.05) {
			frame_time = ms;
			isVideoStopped = !videoCapture.Read(frame!);
			if (!frame!.Empty()) {
				frameBitmap?.Dispose();
				frameBitmap = MatToD2DBitmap(renderTarget!, frame);
				if (!isDetecting) {
					isDetecting = true;
					var image = MatToSxImage(frame, 640, 640);
					_ = RunDetectionAsync(image);
				}
				str_fps = String.Format("FPS: {0:0.00}", fps);
			}
		}

		renderTarget!.BeginDraw();
		RenderScene(ref renderTarget!);
		renderTarget.EndDraw();
#if !USE_HWND_RENDER_TARGET
		swapChain!.Present(0, PresentFlags.None);
#endif
	}

#if !USE_HWND_RENDER_TARGET
	private void RenderScene(ref D2D.RenderTarget target)
#else
	private void RenderScene(ref D2D.WindowRenderTarget target)
#endif
	{
		if (frameBitmap != null) {
			target.DrawBitmap(frameBitmap,
				new RectF(0, 0, ClientSize.Width, ClientSize.Height), 1.0f, BitmapInterpolationMode.Linear);
			if (detection_box.Count > 0) {
				lock (lockBox) {
					// Calculate required scale factory
					var gain = Math.Max(ClientSize.Width, ClientSize.Height) / 640f;
					var aspect = ClientSize.Width / (float)ClientSize.Height;
					string str = "";
					foreach (var box in detection_box) {
						var rc = box.Rect;
						float x = rc.X * gain, y = rc.Y * gain / aspect;
						var rect = new RectF(x, y, x + rc.Width * gain, y + rc.Height * gain / aspect);
						var box_brush = new SolidColorBrush(renderTarget, box.Color);
						target.DrawRectangle(rect, box_brush, 2);
						str = $"{box.Label} " + box.Confidence.ToString("F2");
						var lblSize = MeasureString(target, fontArial20!, str);
						rect.Right = rect.Left + lblSize.Width + 8;
						rect.Bottom = rect.Top + lblSize.Height;
						target.FillRectangle(rect, box_brush);
						box_brush.Dispose();
						rect.Left += 4;
						target.DrawText(str, fontArial20, rect, whiteBrush);
					}
				}
			}
		}
		target.DrawText(str_fps, fontArial30, new RectF(5, ClientSize.Height - 35.0f, 200, ClientSize.Height), greenBrush);
	}

	static async Task RunDetectionAsync(SxImage.Image image)
	{
		var result = await Yolo!.DetectAsync(image);
		isDetecting = false;
		Debug.WriteLine($"Result: {result}");
		Debug.WriteLine($"Speed:  {result.Speed}");
		lock (lockBox) {
			detection_box.Clear();
			if (result.Count > 0) {
				MakeDetectionBox(ref result);
			}
		}
		image.Dispose();
	}

	private record struct DetectionBox
	{
		public Color4 Color;
		public SxImage.Rectangle Rect;
		public string Label;
		public float Confidence;
	}

	const float BoxAlpha = 0.75f;
	static readonly IList<Color4> ColorPalette = new ReadOnlyCollection<Color4>([
		new(0xFF / 255f, 0x38 / 255f, 0x38 / 255f, BoxAlpha),
		new(0xFF / 255f, 0x9D / 255f, 0x97 / 255f, BoxAlpha),
		new(0xFF / 255f, 0x70 / 255f, 0x1F / 255f, BoxAlpha),
		new(0xFF / 255f, 0xB2 / 255f, 0x1D / 255f, BoxAlpha),
		new(0xCF / 255f, 0xD2 / 255f, 0x31 / 255f, BoxAlpha),
		new(0x48 / 255f, 0xF9 / 255f, 0x0A / 255f, BoxAlpha),
		new(0x92 / 255f, 0xCC / 255f, 0x17 / 255f, BoxAlpha),
		new(0x3D / 255f, 0xDB / 255f, 0x86 / 255f, BoxAlpha),
		new(0x1A / 255f, 0x93 / 255f, 0x34 / 255f, BoxAlpha),
		new(0x00 / 255f, 0xD4 / 255f, 0xBB / 255f, BoxAlpha),
		new(0x2C / 255f, 0x99 / 255f, 0xA8 / 255f, BoxAlpha),
		new(0x00 / 255f, 0xC2 / 255f, 0xFF / 255f, BoxAlpha),
		new(0x34 / 255f, 0x45 / 255f, 0x93 / 255f, BoxAlpha),
		new(0x64 / 255f, 0x73 / 255f, 0xFF / 255f, BoxAlpha),
		new(0x00 / 255f, 0x18 / 255f, 0xEC / 255f, BoxAlpha),
		new(0x84 / 255f, 0x38 / 255f, 0xFF / 255f, BoxAlpha),
		new(0x52 / 255f, 0x00 / 255f, 0x85 / 255f, BoxAlpha),
		new(0xCB / 255f, 0x38 / 255f, 0xFF / 255f, BoxAlpha),
		new(0xFF / 255f, 0x95 / 255f, 0xC8 / 255f, BoxAlpha),
		new(0xFF / 255f, 0x37 / 255f, 0xC7 / 255f, BoxAlpha),
	]);

	private static void MakeDetectionBox(ref YoloResult<Detection> result)
	{
		foreach (var box in result) {
			if (box.Confidence < 0.45f)
				continue;
			int color_idx = box.Name.Id % ColorPalette.Count;
			detection_box.Add(new DetectionBox
			{
				Color = ColorPalette[color_idx],
				Rect = box.Bounds,
				Label = box.Name.ToString().Split('\'')[1].Trim('\''),
				Confidence = box.Confidence
			});
		}
	}

	private static D2D.Factory? d2dFactory;
	private static DWrite.Factory? dwriteFactory;
#if !USE_HWND_RENDER_TARGET
	private static DXDevice? device;
	private SwapChain? swapChain;
	private Texture2D? backBuffer;
	private Surface? surface;
	private RenderTargetView? renderView;
	private RenderTarget? renderTarget;
#else
	private WindowRenderTarget? renderTarget;
#endif

	private VideoCapture? videoCapture;
	private Mat? frame;
	private D2D.Bitmap? frameBitmap;
	private static List<DetectionBox> detection_box = [];
	private bool isVideoStopped = true;

	private SolidColorBrush? whiteBrush;
	private SolidColorBrush? greenBrush;
	private DWrite.TextFormat? fontArial20;
	private DWrite.TextFormat? fontArial30;

	private const int WM_ERASEBKGND = 0x0014;
	private const int WM_SETTINGCHANGE = 0x001A;

	protected override void WndProc(ref System.Windows.Forms.Message m)
	{
		switch (m.Msg) {
			case WM_ERASEBKGND:
				if (isVideoStopped) {
					if (frameBitmap != null) {
						renderTarget!.BeginDraw();
						RenderScene(ref renderTarget!);
						renderTarget.EndDraw();
#if !USE_HWND_RENDER_TARGET
						swapChain!.Present(0, PresentFlags.None);
#endif
					} else {
						using var g = CreateGraphics();
						g.Clear(System.Drawing.Color.Black);
					}
				}
				break;

			case WM_SETTINGCHANGE:
				isDarkMode = !DarkMode.IsLightTheme();
				DarkMode.UseImmersiveDarkMode(Handle, isDarkMode);
				break;

			default:
				base.WndProc(ref m);
				break;
		}
	}

	private static BitmapProperties defaultBitmapProp = new(new PixelFormat(Format.B8G8R8A8_UNorm, D2D.AlphaMode.Premultiplied));

	private static D2D.Bitmap MatToD2DBitmap(D2D.RenderTarget rt, Mat mat)
	{
		using Mat image = mat.CvtColor(ColorConversionCodes.BGR2BGRA);
		DataPointer dataPointer = new(image.Data, (int)image.Total() * image.ElemSize());
		return new D2D.Bitmap(rt, new Size2(image.Width, image.Height), dataPointer, image.Width * Unsafe.SizeOf<Bgra32>(), defaultBitmapProp);
	}

	private static D2D.Bitmap SxImageToBitmap(D2D.RenderTarget rt, SxImage.Image src)
	{
		var image = src.CloneAs<Bgra32>();
		int widthbytes = image.Width * Unsafe.SizeOf<Bgra32>();
		if (!image.DangerousTryGetSinglePixelMemory(out Memory<Bgra32> memory)) {
			throw new Exception(
				"This can only happen with multi-GB images or when PreferContiguousImageBuffers is not set to true.");
		}
		unsafe {
			using MemoryHandle pinHandle = memory.Pin();
			DataPointer dataPointer = new(pinHandle.Pointer, widthbytes * image.Height);
			return new D2D.Bitmap(rt, new Size2(image.Width, image.Height), dataPointer, widthbytes, defaultBitmapProp);
		}
	}

	private static SxImage.Image<Bgr24> MatToSxImage(Mat mat, int w, int h)
	{
		Mat resized = new();
		Cv2.Resize(mat, resized, new CvSize(w, h));
		unsafe {
			ReadOnlySpan<byte> ptr = new((void*)resized.Data, resized.Width * resized.Height * Unsafe.SizeOf<Bgr24>());
			return SxImage.Image.LoadPixelData<Bgr24>(ptr, resized.Width, resized.Height);
		}
	}

	public static Size2F MeasureString(D2D.RenderTarget RenderTarget, DWrite.TextFormat textFormat,	string text)
	{
		float width = RenderTarget.Size.Width * 2;
		using DWrite.TextLayout layout = new(dwriteFactory, text, textFormat, width, textFormat.FontSize);
		return new Size2F(layout.Metrics.Width - layout.Metrics.Left, layout.Metrics.Height - layout.Metrics.Top);
	}
}
