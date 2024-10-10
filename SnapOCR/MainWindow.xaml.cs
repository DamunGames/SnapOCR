using Microsoft.VisualBasic;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace SnapOCR
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private const int CaptureInterval = 100;

		private const int HOTKEY_ID_REGISTER = 9000;
		private const int HOTKEY_ID_DELETE = 9001;
		private const uint MOD_CTRL = 0x0002; // Ctrlキー
		private const uint VK_BACK_SLASH = 0xE2; // BackSlashキー
		private const uint VK_DELETE = 0x2E; // Deleteキー

		// Win32 APIを使用するためのDllImport
		[DllImport("user32.dll")]
		private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

		[DllImport("user32.dll")]
		private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

		public class TextData
		{
			public string Text { get; set; }
		}
		public ObservableCollection<TextData> RegisterdTexts { get; set; } = new ObservableCollection<TextData>();

		private int _timer;
		private CaptureRectWindow? _captureRectWindow = null;

		// ==========================
		// MainWindow
		// ==========================
		public MainWindow()
		{
			InitializeComponent();

			DataContext = this;

			Loaded += Window_Loaded;
			Closed += Window_Closed;
		}

		private void MainWindow_Rendering(object? sender, EventArgs e)
		{
			_timer++;
			if (_timer >= CaptureInterval) {
				SnapAndOCR();
				_timer = 0;
			}
		}

		// ==========================
		// CaptureRectWindow
		// ==========================
		private void OpenCaptureRectWindow()
		{
			if (_captureRectWindow == null) {
				_captureRectWindow = new CaptureRectWindow();
				_captureRectWindow.Closing += CaptureRectWindow_Closing;
				_captureRectWindow.Show();
			}
		}

		private void CaptureRectWindow_Closing(object? sender, CancelEventArgs e)
		{
			Close();
		}

		// ==========================
		// TextData
		// ==========================
		private void RegisterTextData(string text)
		{
			string[] splited = Regex.Replace(text, "\r\n", "\n").Split(new char[] { '\n', '\r' });
			for (int i = 0; i < splited.Length; i++) {
				RegisterdTexts.Add(new TextData()
				{
					Text = splited[i],
				});
			}
		}

		private void RemoveTextData()
		{
			if (RegisterdTexts.Count > 0) {
				RegisterdTexts.RemoveAt(RegisterdTexts.Count - 1);
			}
		}

		private void SaveRegisterdTexts()
		{
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = $"SnapOCR_{DateTime.Now:yyyy_MM_dd_HH_mm_ss}.txt";
			saveFileDialog.Filter = "すべてのファイル|*.*|テキスト ファイル|*.txt";
			saveFileDialog.FilterIndex = 2;
			if (saveFileDialog.ShowDialog() == true) {
				string path = saveFileDialog.FileName;
				List<string> lines = new List<string>(RegisterdTexts.Count);
				foreach (var ocrResult in RegisterdTexts) {
					lines.Add(ocrResult.Text);
				}
				File.WriteAllLines(path, lines);
			}
		}

		// ==========================
		// Bitmap
		// ==========================
		private Bitmap? SnapCaptureRectWindow()
		{
			if (_captureRectWindow == null) return null;
			System.Windows.Point capturePoint = _captureRectWindow.PointToScreen(new System.Windows.Point());
			System.Windows.Size captureSize = _captureRectWindow.RenderSize;
			Bitmap bitmap = CaptureScreen(new System.Drawing.Rectangle((int)capturePoint.X + 1, (int)capturePoint.Y + 1, (int)captureSize.Width - 2, (int)captureSize.Height - 2));
			return AddMargin(bitmap, 30, EstimateBackgroundColor(bitmap));
		}

		private static Bitmap CaptureScreen(System.Drawing.Rectangle rect)
		{
			// 指定された矩形の領域をキャプチャする
			Bitmap bitmap = new Bitmap(rect.Width, rect.Height);
			using (Graphics g = Graphics.FromImage(bitmap)) {
				g.CopyFromScreen(rect.X, rect.Y, 0, 0, rect.Size, CopyPixelOperation.SourceCopy);
			}
			return bitmap;
		}

		private static Bitmap AddMargin(Bitmap originalBitmap, int marginSize, System.Drawing.Color marginColor)
		{
			// 元のビットマップのサイズを取得
			int originalWidth = originalBitmap.Width;
			int originalHeight = originalBitmap.Height;

			// 新しいビットマップのサイズを計算
			int newWidth = originalWidth + marginSize * 2;
			int newHeight = originalHeight + marginSize * 2;

			// 新しいビットマップを作成（指定した背景色で塗りつぶし）
			Bitmap newBitmap = new Bitmap(newWidth, newHeight);
			using (Graphics g = Graphics.FromImage(newBitmap)) {
				// 新しいビットマップ全体をマージン色で塗りつぶす
				g.Clear(marginColor);

				// 元のビットマップを中央に描画
				g.DrawImage(originalBitmap, marginSize, marginSize, originalWidth, originalHeight);
			}

			return newBitmap;
		}

		private static System.Drawing.Color EstimateBackgroundColor(Bitmap bitmap, int sampleStep = 10)
		{
			// 色をカウントするための辞書
			Dictionary<System.Drawing.Color, int> colorCount = new Dictionary<System.Drawing.Color, int>();

			// サンプリング間隔 (デフォルトは10ピクセルごとにサンプリング)
			int width = bitmap.Width;
			int height = bitmap.Height;

			// 画像の四隅をサンプリング
			for (int x = 0; x < width; x += sampleStep) {
				for (int y = 0; y < height; y += sampleStep) {
					// ピクセルの色を取得
					System.Drawing.Color pixelColor = bitmap.GetPixel(x, y);

					// 辞書に色をカウント
					if (colorCount.ContainsKey(pixelColor)) {
						colorCount[pixelColor]++;
					}
					else {
						colorCount[pixelColor] = 1;
					}
				}
			}

			// 最も多く出現した色を背景色として返す
			return colorCount.OrderByDescending(c => c.Value).First().Key;
		}

		private static SoftwareBitmap ConvertBitmapToSoftwareBitmap(Bitmap bitmap)
		{
			// Bitmapをメモリストリームに書き込み
			using (var stream = new MemoryStream()) {
				bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
				stream.Seek(0, SeekOrigin.Begin);

				// ストリームをIRandomAccessStreamに変換
				var randomAccessStream = new InMemoryRandomAccessStream();
				var outputStream = randomAccessStream.AsStreamForWrite();
				stream.CopyTo(outputStream);
				outputStream.Flush();
				randomAccessStream.Seek(0);

				// BitmapDecoderでSoftwareBitmapを生成
				var decoder = Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(randomAccessStream).AsTask().Result;
				var softwareBitmap = decoder.GetSoftwareBitmapAsync().AsTask().Result;

				return softwareBitmap;
			}
		}

		// ==========================
		// OCR
		// ==========================
		private void SnapAndOCR()
		{
			Bitmap? bitmap = SnapCaptureRectWindow();
			if (bitmap == null) {
				Console.WriteLine("Failed");
				OCRResultTextBlock.Text = "Failed";
				return;
			}

			SoftwareBitmap softwareBitmap = ConvertBitmapToSoftwareBitmap(bitmap);

			// OCRエンジンを初期化
			OcrEngine ocrEngine = OcrEngine.TryCreateFromUserProfileLanguages();
			OcrResult ocrResult = ocrEngine.RecognizeAsync(softwareBitmap).AsTask().Result;

			if (ocrResult.Lines != null && ocrResult.Lines.Count > 0) {
				Console.WriteLine($"Sucess Lines:{ocrResult.Lines.Count}");

				var stringBuilder = new StringBuilder();
				List<string> results = new List<string>();

				string text = Regex.Replace(ocrResult.Lines[0].Text, " ", "");
				stringBuilder.Append(text);
				results.Add(text);
				Console.WriteLine(text);
				for (int i = 1; i < ocrResult.Lines.Count; i++) {
					text = Regex.Replace(ocrResult.Lines[i].Text, " ", "");
					stringBuilder.AppendLine();
					stringBuilder.Append(text);
					results.Add(text);
					Console.WriteLine(text);
				}
				OCRResultTextBlock.Text = stringBuilder.ToString();
			}
			else {
				Console.WriteLine("Failed");
				OCRResultTextBlock.Text = "Failed";
			}
		}

		// ==========================
		// HotKey
		// ==========================
		private void AddHook()
		{
			var hwndSource = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
			hwndSource.AddHook(HwndHook);
		}

		private void RegisterHotKeys()
		{
			// グローバルホットキーの登録 (Ctrl + VK_?)
			RegisterHotKey(new WindowInteropHelper(this).Handle, HOTKEY_ID_REGISTER, MOD_CTRL, VK_BACK_SLASH);
			RegisterHotKey(new WindowInteropHelper(this).Handle, HOTKEY_ID_DELETE, MOD_CTRL, VK_DELETE);
		}

		private void UnregisterHotKeys()
		{
			// グローバルホットキーの登録解除
			UnregisterHotKey(new WindowInteropHelper(this).Handle, HOTKEY_ID_REGISTER);
			UnregisterHotKey(new WindowInteropHelper(this).Handle, HOTKEY_ID_DELETE);
		}

		private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
		{
			const int WM_HOTKEY = 0x0312;
			if (msg == WM_HOTKEY) {
				int keyId = wParam.ToInt32();
				if (keyId == HOTKEY_ID_REGISTER) {
					// Ctrl + \が押された時の処理
					RegisterTextData(OCRResultTextBlock.Text);

					handled = true;
				}
				else if (keyId == HOTKEY_ID_DELETE) {
					// Ctrl + Delete
					RemoveTextData();

					handled = true;
				}
			}
			return IntPtr.Zero;

		}

		// ==========================
		// ウィンドウイベント
		// ==========================
		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			CompositionTarget.Rendering += MainWindow_Rendering;

			AddHook();
			RegisterHotKeys();

			OpenCaptureRectWindow();
		}

		private void Window_Closed(object? sender, EventArgs e)
		{
			CompositionTarget.Rendering -= MainWindow_Rendering;

			UnregisterHotKeys();
		}

		// ==========================
		// ボタンイベント
		// ==========================
		private void RegisterButton_Click(object sender, RoutedEventArgs e)
		{
			RegisterTextData(OCRResultTextBlock.Text);
		}

		private void RemoveRowButton_Click(object sender, RoutedEventArgs e)
		{
			RemoveTextData();
		}

		private void SaveButton_Click(object sender, RoutedEventArgs e)
		{
			SaveRegisterdTexts();
		}
	}
}