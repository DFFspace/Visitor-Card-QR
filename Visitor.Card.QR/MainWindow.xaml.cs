using Microsoft.UI;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using QRCoder;
using System;
using System.Globalization;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using WinRT.Interop;
using WinUIEx;

namespace Visitor.Card.QR;

public sealed partial class MainWindow : WindowEx
{
    private byte[]? _lastPngBytes;
    private readonly DispatcherQueueTimer _debounceTimer;
    private readonly AppWindow appWindow;

    public MainWindow()
    {
        InitializeComponent();
        SystemBackdrop = new MicaBackdrop();
        ExtendsContentIntoTitleBar = true;
        appWindow = GetAppWindowForCurrentWindow();
        appWindow.Resize(new Windows.Graphics.SizeInt32(1100, 1000));

        _debounceTimer = DispatcherQueue.CreateTimer();
        _debounceTimer.Interval = TimeSpan.FromMilliseconds(150);
        _debounceTimer.Tick += (_, __) =>
        {
            _debounceTimer.Stop();
            UpdatePreviewAsync();
        };

        FirstNameBox.Text = "";
        LastNameBox.Text = "";
        FullNameBox.Text = "";
        FunctionBox.Text = "";
        PhoneBox.Text = "+316";
        EmailBox.Text = "@example.nl";
        WebsiteBox.Text = "https://www.example.nl/";
    }

    private AppWindow GetAppWindowForCurrentWindow()
    {
        IntPtr hWnd = WindowNative.GetWindowHandle(this);
        WindowId myWndId = Win32Interop.GetWindowIdFromWindow(hWnd);
        return AppWindow.GetFromWindowId(myWndId);
    }

    private void AnyFieldChanged(object sender, TextChangedEventArgs e) => _debounceTimer?.Restart();
    private void OnEccChanged(object sender, SelectionChangedEventArgs e) => _debounceTimer?.Restart();
    private void OnSizeChanged(object sender, RangeBaseValueChangedEventArgs e) => _debounceTimer?.Restart();

    private void OnTransparentToggled(object sender, RoutedEventArgs e)
    {
        LightHexBox.IsEnabled = !(TransparentBgCheck.IsChecked ?? false);
        _debounceTimer?.Restart();
    }

    private static string BuildVCard(string last, string first, string full, string function, string phone, string email, string website)
    {
        string NL = "\r\n";

        var v = $"BEGIN:VCARD{NL}" +
                $"VERSION:3.0{NL}" +
                $"N:{last};{first}{NL}" +
                $"FN:{full}{NL}" +
                $"ORG:COMPANYNAME{NL}" +
                $"TITLE:{function}{NL}" +
                $"ADR:;;STREET+HOUSENUMB;CITY;STATE;POSTALCODE;COUNTRY{NL}" +
                $"TEL;WORK;VOICE:{NL}" +
                $"TEL;CELL:{phone}{NL}" +
                $"TEL;FAX:{NL}" +
                $"EMAIL;WORK;INTERNET:{email}{NL}" +
                $"URL:{website}{NL}" +
                $"END:VCARD{NL}";

        return v;
    }

    private QRCodeGenerator.ECCLevel GetEcc()
    {
        var tag = "H";
        return tag switch
        {
            "L" => QRCodeGenerator.ECCLevel.L,
            "M" => QRCodeGenerator.ECCLevel.M,
            "Q" => QRCodeGenerator.ECCLevel.Q,
            "H" => QRCodeGenerator.ECCLevel.H,
            _ => QRCodeGenerator.ECCLevel.M
        };
    }

    private static bool TryParseHexColor(string hex, out byte r, out byte g, out byte b, out byte a)
    {
        r = g = b = 0; a = 255;
        if (string.IsNullOrWhiteSpace(hex)) return false;

        hex = hex.Trim();
        if (hex.StartsWith("#")) hex = hex[1..];

        if (hex.Length is 6 or 8)
        {
            try
            {
                int idx = 0;
                if (hex.Length == 8)
                {
                    a = byte.Parse(hex.Substring(idx, 2), NumberStyles.HexNumber);
                    idx += 2;
                }
                r = byte.Parse(hex.Substring(idx, 2), NumberStyles.HexNumber);
                g = byte.Parse(hex.Substring(idx + 2, 2), NumberStyles.HexNumber);
                b = byte.Parse(hex.Substring(idx + 4, 2), NumberStyles.HexNumber);
                return true;
            }
            catch { }
        }
        return false;
    }

    private async System.Threading.Tasks.Task UpdatePreviewAsync()
    {
        var last = LastNameBox.Text?.Trim() ?? "";
        var first = FirstNameBox.Text?.Trim() ?? "";
        var full = FullNameBox.Text?.Trim() ?? "";
        var role = FunctionBox.Text?.Trim() ?? "";
        var phone = PhoneBox.Text?.Trim() ?? "";
        var email = EmailBox.Text?.Trim() ?? "";
        var site = WebsiteBox.Text?.Trim() ?? "";

        var vcard = BuildVCard(last, first, full, role, phone, email, site);

        if (!TryParseHexColor(DarkHexBox.Text, out var dr, out var dg, out var db, out var da))
        {
            dr = 0; dg = 0; db = 0; da = 255;
        }

        byte lr = 255, lg = 255, lb = 255, la = 255;
        bool transparent = TransparentBgCheck.IsChecked == true;
        if (!transparent)
        {
            if (!TryParseHexColor(LightHexBox.Text, out lr, out lg, out lb, out la))
            {
                lr = 255; lg = 255; lb = 255; la = 255;
            }
        }
        else
        {
            la = 0;
        }

        int targetSize = (int)SizeSlider.Value;
        var ecc = GetEcc();

        using var gen = new QRCodeGenerator();
        using var data = gen.CreateQrCode(vcard, ecc);

        int modules = data.ModuleMatrix.Count;
        int ppm = Math.Max(1, targetSize / modules);

        var png = new PngByteQRCode(data);
        byte[] bytes = png.GetGraphic(
            pixelsPerModule: ppm,
            darkColorRgba: new byte[] { dr, dg, db, da },
            lightColorRgba: new byte[] { lr, lg, lb, la },
            drawQuietZones: true
        );

        _lastPngBytes = bytes;

        using InMemoryRandomAccessStream stream = new();
        await stream.WriteAsync(bytes.AsBuffer());
        stream.Seek(0);

        BitmapImage bmp = new();
        await bmp.SetSourceAsync(stream);
        PreviewImage.Source = bmp;
        PreviewImage.Width = bmp.PixelWidth;
        PreviewImage.Height = bmp.PixelHeight;
    }

    private async void OnSaveClicked(object sender, RoutedEventArgs e)
    {
        if (_lastPngBytes is null || _lastPngBytes.Length == 0)
        {
            return;
        }

        var picker = new FileSavePicker
        {
            SuggestedStartLocation = PickerLocationId.PicturesLibrary,
            SuggestedFileName = string.IsNullOrWhiteSpace(FullNameBox.Text)
                ? "qr-code"
                : $"{FullNameBox.Text.Trim()}-qr"
        };
        picker.FileTypeChoices.Add("PNG image", new[] { ".png" });

        IntPtr hWnd = WindowNative.GetWindowHandle(this);
        InitializeWithWindow.Initialize(picker, hWnd);

        StorageFile file = await picker.PickSaveFileAsync();
        if (file is null) return;

        await FileIO.WriteBytesAsync(file, _lastPngBytes);
    }
}

internal static class TimerExtensions
{
    public static void Restart(this DispatcherQueueTimer? timer)
    {
        if (timer is null) return;
        if (timer.IsRunning) timer.Stop();
        timer.Start();
    }
}
