using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Foundation;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace WinUI3TemplateEditor;

public sealed partial class MainPage : Page
{
    private const string DefaultOutputDir = @"C:\Users\lauta\Downloads\docs\docs\GUIAS DE CAMBIOS\EXPORTADOS";
    private const string DefaultUpdateUrl = "https://api.github.com/repos/lauty23/editor-de-plantillas/releases/latest";
    private const string DefaultReleaseUrl = "https://github.com/lauty23/editor-de-plantillas/releases/latest";
    private static readonly string CurrentAppVersion = GetCurrentAppVersion();
    private const double SidebarExpandedWidth = 300;
    private const double SidebarCollapsedWidth = 48;
    private static readonly TimeSpan SidebarAnimationDuration = TimeSpan.FromMilliseconds(220);
    private readonly ObservableCollection<TemplateProfile> _profiles = new();
    private readonly List<EditableField> _fields = new();
    private readonly List<BallField> _ballFields = new();
    private readonly Stopwatch _sidebarAnimationClock = new();
    private TemplateProfile? _selectedProfile;
    private string _outputDir = DefaultOutputDir;
    private string _updateUrl = DefaultUpdateUrl;
    private double _zoom = 1.0;
    private string? _previewImagePath;
    private string? _signatureImagePath;
    private bool _sidebarVisible = true;
    private bool _sidebarAnimationRunning;
    private bool _autoUpdateChecked;
    private double _sidebarAnimationStart;
    private double _sidebarAnimationTarget;

    public MainPage()
    {
        InitializeComponent();
        DocumentsCombo.ItemsSource = _profiles;
        LoadConfig();
        if (_profiles.Count > 0)
        {
            DocumentsCombo.SelectedIndex = 0;
        }
        UpdateZoom();
        UpdateSidebarState();
        SidebarPanel.SizeChanged += (_, _) => UpdateSidebarClip(SidebarColumn.Width.Value);
        Loaded += MainPage_Loaded;
    }

    private async void MainPage_Loaded(object sender, RoutedEventArgs e)
    {
        if (_autoUpdateChecked)
        {
            return;
        }

        _autoUpdateChecked = true;
        await CheckForUpdatesAsync(manual: false);
    }

    private void LoadConfig()
    {
        try
        {
            var configPath = FindConfigPath();
            var json = File.ReadAllText(configPath);
            var config = JsonSerializer.Deserialize<AppConfig>(json, JsonOptions()) ?? new AppConfig();
            _outputDir = string.IsNullOrWhiteSpace(config.OutputDir) ? DefaultOutputDir : config.OutputDir;
            _updateUrl = string.IsNullOrWhiteSpace(config.UpdateUrl) ? DefaultUpdateUrl : config.UpdateUrl;
            _profiles.Clear();
            foreach (var document in config.Documents)
            {
                _profiles.Add(document);
            }
            Status("Configuracion cargada.", InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"No se pudo cargar config/edit_profiles.json: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private static string FindConfigPath()
    {
        var basePath = AppContext.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(basePath, "config", "edit_profiles.json"),
            Path.GetFullPath(Path.Combine(basePath, "..", "..", "..", "..", "..", "config", "edit_profiles.json")),
            Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "config", "edit_profiles.json"))
        };
        return candidates.First(File.Exists);
    }

    private static JsonSerializerOptions JsonOptions() => new()
    {
        PropertyNameCaseInsensitive = true
    };

    private void DocumentsCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DocumentsCombo.SelectedItem is TemplateProfile profile)
        {
            LoadProfile(profile);
        }
    }

    private void LoadProfile(TemplateProfile profile)
    {
        _selectedProfile = profile;
        _fields.Clear();
        _ballFields.Clear();
        FieldsPanel.Children.Clear();
        PreviewSubtitle.Text = $"{profile.Label} · {profile.Kind.ToUpperInvariant()}";

        if (!File.Exists(profile.SourcePath))
        {
            Status($"No existe el archivo: {profile.SourcePath}", InfoBarSeverity.Error);
            RenderPreview();
            return;
        }

        try
        {
            if (profile.Kind.Equals("docx", StringComparison.OrdinalIgnoreCase))
            {
                _fields.AddRange(DocxTemplate.ReadFields(profile.SourcePath));
                _ballFields.AddRange(DocxTemplate.ReadBallFields(profile.SourcePath));
                RenderFieldEditors();
                Status($"{_fields.Count} campos y {_ballFields.Sum(item => item.Values.Count)} numeros de balotas detectados.", InfoBarSeverity.Success);
            }
            else if (profile.Kind.Equals("png", StringComparison.OrdinalIgnoreCase))
            {
                Status("PNG cargado. Esta version WinUI exporta la imagen base; la edicion por zonas sigue en el ejecutable anterior.", InfoBarSeverity.Warning);
            }
            else
            {
                Status("Formato no reconocido. Se exportara una copia base.", InfoBarSeverity.Warning);
            }
        }
        catch (Exception ex)
        {
            Status($"No se pudo abrir la plantilla: {ex.Message}", InfoBarSeverity.Error);
        }

        RenderPreview();
    }

    private void RenderFieldEditors()
    {
        FieldsPanel.Children.Clear();
        if (_fields.Count == 0)
        {
            if (_ballFields.Count > 0)
            {
                RenderBallEditors();
                return;
            }
            {
                FieldsPanel.Children.Add(new TextBlock
                {
                    Text = "No se detectaron zonas amarillas editables.",
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
                });
                return;
            }
        }

        RenderBallEditors();
        RenderSignatureEditor();

        foreach (var field in _fields)
        {
            var textBox = new TextBox
            {
                Text = field.Value,
                AcceptsReturn = field.Value.Length > 42,
                TextWrapping = TextWrapping.Wrap,
                MinHeight = field.Value.Length > 42 ? 84 : 44
            };
            textBox.TextChanged += (_, _) =>
            {
                field.Value = textBox.Text;
                RenderPreview();
            };

            var panel = new StackPanel { Spacing = 8 };
            panel.Children.Add(new TextBlock
            {
                Text = field.Label,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
            });
            panel.Children.Add(new TextBlock
            {
                Text = Shorten(field.Original, 70),
                TextWrapping = TextWrapping.Wrap,
                Foreground = (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
            });
            panel.Children.Add(textBox);

            FieldsPanel.Children.Add(new Border
            {
                Padding = new Thickness(12),
                CornerRadius = new CornerRadius(8),
                BorderBrush = (Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
                BorderThickness = new Thickness(1),
                Background = (Brush)Application.Current.Resources["CardBackgroundFillColorDefaultBrush"],
                Child = panel
            });
        }
    }

    private void RenderSignatureEditor()
    {
        var body = new StackPanel { Spacing = 10 };
        body.Children.Add(new TextBlock
        {
            Text = "Firma",
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
        });
        body.Children.Add(new TextBlock
        {
            Text = _signatureImagePath is null
                ? "Usando la firma original del documento."
                : $"Firma seleccionada: {Path.GetFileName(_signatureImagePath)}",
            TextWrapping = TextWrapping.Wrap,
            Foreground = (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
        });

        var buttons = new Grid
        {
            ColumnSpacing = 8
        };
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        buttons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var selectButton = new Button
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Content = "Elegir imagen"
        };
        selectButton.Click += SelectSignatureButton_Click;
        buttons.Children.Add(selectButton);

        var resetButton = new Button
        {
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Content = "Original",
            IsEnabled = _signatureImagePath is not null
        };
        resetButton.Click += (_, _) =>
        {
            _signatureImagePath = null;
            RenderFieldEditors();
            RenderPreview();
        };
        Grid.SetColumn(resetButton, 1);
        buttons.Children.Add(resetButton);

        body.Children.Add(buttons);

        FieldsPanel.Children.Add(new Border
        {
            Padding = new Thickness(12),
            CornerRadius = new CornerRadius(8),
            BorderBrush = (Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
            BorderThickness = new Thickness(1),
            Background = (Brush)Application.Current.Resources["CardBackgroundFillColorDefaultBrush"],
            Child = body
        });
    }

    private async void SelectSignatureButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".png");
            picker.FileTypeFilter.Add(".jpg");
            picker.FileTypeFilter.Add(".jpeg");
            picker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;

            if (Application.Current is App app && app.MainWindow is not null)
            {
                InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(app.MainWindow));
            }

            var file = await picker.PickSingleFileAsync();
            if (file is null)
            {
                return;
            }

            _signatureImagePath = file.Path;
            RenderFieldEditors();
            RenderPreview();
            Status($"Firma cargada: {file.Name}", InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"No se pudo cargar la firma: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private void RenderBallEditors()
    {
        foreach (var ballField in _ballFields)
        {
            var body = new StackPanel { Spacing = 10 };
            body.Children.Add(new TextBlock
            {
                Text = ballField.Label,
                FontWeight = Microsoft.UI.Text.FontWeights.SemiBold
            });
            body.Children.Add(new TextBlock
            {
                Text = "Edita los numeros y se exportaran con formato de balota.",
                TextWrapping = TextWrapping.Wrap,
                Foreground = (Brush)Application.Current.Resources["TextFillColorSecondaryBrush"]
            });

            var grid = new Grid
            {
                ColumnSpacing = 8,
                RowSpacing = 8
            };
            for (var column = 0; column < 3; column++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            }
            for (var rowIndex = 0; rowIndex < 2; rowIndex++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }

            for (var index = 0; index < ballField.Values.Count; index++)
            {
                var valueIndex = index;
                var box = new TextBox
                {
                    Text = ballField.Values[index],
                    MaxLength = 2,
                    MinHeight = 36,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    HorizontalContentAlignment = HorizontalAlignment.Center
                };
                box.TextChanged += (_, _) =>
                {
                    ballField.Values[valueIndex] = SanitizeBallValue(box.Text);
                    RenderPreview();
                };
                Grid.SetColumn(box, index % 3);
                Grid.SetRow(box, index / 3);
                grid.Children.Add(box);
            }
            body.Children.Add(grid);

            FieldsPanel.Children.Add(new Border
            {
                Padding = new Thickness(12),
                CornerRadius = new CornerRadius(8),
                BorderBrush = (Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
                BorderThickness = new Thickness(1),
                Background = (Brush)Application.Current.Resources["CardBackgroundFillColorDefaultBrush"],
                Child = body
            });
        }
    }

    private void RenderPreview()
    {
        PreviewContent.Children.Clear();
        PreviewImage.Source = null;
        PreviewImage.Visibility = Visibility.Collapsed;
        PreviewContent.Visibility = Visibility.Visible;

        if (_selectedProfile is not null && File.Exists(_selectedProfile.SourcePath))
        {
            try
            {
                if (_selectedProfile.Kind.Equals("docx", StringComparison.OrdinalIgnoreCase))
                {
                    _previewImagePath = Path.Combine(Path.GetTempPath(), $"winui_template_preview_{Guid.NewGuid():N}.png");
                    DocxPreviewRenderer.Render(_selectedProfile.SourcePath, _previewImagePath, _fields, _ballFields, _signatureImagePath);
                    PreviewImage.Source = new BitmapImage(new Uri(_previewImagePath));
                    PreviewImage.Visibility = Visibility.Visible;
                    PreviewContent.Visibility = Visibility.Collapsed;
                    return;
                }
                if (_selectedProfile.Kind.Equals("png", StringComparison.OrdinalIgnoreCase))
                {
                    PreviewImage.Source = new BitmapImage(new Uri(_selectedProfile.SourcePath));
                    PreviewImage.Visibility = Visibility.Visible;
                    PreviewContent.Visibility = Visibility.Collapsed;
                    return;
                }
            }
            catch
            {
                PreviewImage.Visibility = Visibility.Collapsed;
                PreviewContent.Visibility = Visibility.Visible;
            }
        }

        var titleGrid = new Grid
        {
            BorderBrush = new SolidColorBrush(Colors.Black),
            BorderThickness = new Thickness(1),
            MinHeight = 38
        };
        titleGrid.ColumnDefinitions.Add(new ColumnDefinition());
        titleGrid.ColumnDefinitions.Add(new ColumnDefinition());
        titleGrid.Children.Add(new TextBlock
        {
            Text = _selectedProfile?.Label ?? "PLANTILLA",
            Foreground = new SolidColorBrush(Colors.Black),
            FontWeight = Microsoft.UI.Text.FontWeights.Bold,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(12, 0, 0, 0)
        });
        var rightTitle = new TextBlock
        {
            Text = "Campos editables",
            Foreground = new SolidColorBrush(Colors.Black),
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 12, 0)
        };
        Grid.SetColumn(rightTitle, 1);
        titleGrid.Children.Add(rightTitle);
        PreviewContent.Children.Add(titleGrid);

        if (_selectedProfile is null)
        {
            PreviewContent.Children.Add(PreviewText("Selecciona una plantilla para verla aqui.", 16, true));
            return;
        }

        if (_fields.Count == 0)
        {
            PreviewContent.Children.Add(PreviewText("La plantilla seleccionada no tiene campos amarillos detectados en esta version WinUI.", 14, false));
            return;
        }

        foreach (var field in _fields)
        {
            PreviewContent.Children.Add(PreviewText(field.Label, 11, true));
            PreviewContent.Children.Add(PreviewText(SpanishProofing.Correct(field.Value), 14, false));
        }

        foreach (var ballField in _ballFields)
        {
            PreviewContent.Children.Add(PreviewText(ballField.Label, 11, true));
            PreviewContent.Children.Add(PreviewText(string.Join("  ", ballField.Values), 18, true));
        }
    }

    private static TextBlock PreviewText(string text, double size, bool bold) => new()
    {
        Text = text,
        TextWrapping = TextWrapping.Wrap,
        Foreground = new SolidColorBrush(Colors.Black),
        FontSize = size,
        FontWeight = bold ? Microsoft.UI.Text.FontWeights.SemiBold : Microsoft.UI.Text.FontWeights.Normal
    };

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedProfile is null)
        {
            Status("Selecciona un documento primero.", InfoBarSeverity.Warning);
            return;
        }

        try
        {
            Directory.CreateDirectory(_outputDir);
            var destination = Path.Combine(_outputDir, ForcedOutputName(_selectedProfile));
            if (_selectedProfile.Kind.Equals("docx", StringComparison.OrdinalIgnoreCase))
            {
                PdfExport.FromDocx(_selectedProfile.SourcePath, destination, _fields, _ballFields, _signatureImagePath);
            }
            else if (_selectedProfile.Kind.Equals("png", StringComparison.OrdinalIgnoreCase)
                     || _selectedProfile.Kind.Equals("jpg", StringComparison.OrdinalIgnoreCase)
                     || _selectedProfile.Kind.Equals("jpeg", StringComparison.OrdinalIgnoreCase))
            {
                PdfExport.FromImage(_selectedProfile.SourcePath, destination);
            }
            else
            {
                throw new NotSupportedException("Solo se puede exportar DOCX o imagen como PDF.");
            }

            Status($"Exportado: {destination}", InfoBarSeverity.Success);
            await new ContentDialog
            {
                XamlRoot = XamlRoot,
                Title = "Archivo exportado",
                Content = destination,
                CloseButtonText = "Listo"
            }.ShowAsync();
        }
        catch (Exception ex)
        {
            Status($"No se pudo exportar: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        Directory.CreateDirectory(_outputDir);
        Process.Start(new ProcessStartInfo
        {
            FileName = _outputDir,
            UseShellExecute = true
        });
    }

    private async void CheckUpdatesButton_Click(object sender, RoutedEventArgs e)
    {
        await CheckForUpdatesAsync(manual: true);
    }

    private async Task CheckForUpdatesAsync(bool manual)
    {
        if (string.IsNullOrWhiteSpace(_updateUrl))
        {
            if (manual)
            {
                Status("No hay una URL de actualizacion configurada.", InfoBarSeverity.Warning);
            }
            return;
        }

        try
        {
            var updateUri = new Uri(_updateUrl);
            if (!SupportsAutomaticUpdate(updateUri))
            {
                if (manual)
                {
                    Status("Abriendo pagina de actualizaciones.", InfoBarSeverity.Informational);
                    await Windows.System.Launcher.LaunchUriAsync(updateUri);
                }
                return;
            }

            if (manual)
            {
                Status("Buscando actualizaciones...", InfoBarSeverity.Informational);
            }

            var manifest = await FetchUpdateManifestAsync(updateUri);
            if (!IsNewerVersion(manifest.Version, CurrentAppVersion))
            {
                if (manual)
                {
                    Status($"Ya tienes la ultima version ({CurrentAppVersion}).", InfoBarSeverity.Success);
                }
                return;
            }

            var downloadUrl = FirstNonEmpty(manifest.DownloadUrl);
            if (string.IsNullOrWhiteSpace(downloadUrl))
            {
                await ShowReleaseFallbackAsync(manifest, manual);
                return;
            }

            Status($"Descargando actualizacion {manifest.Version}...", InfoBarSeverity.Informational);
            var installerPath = await DownloadUpdateAsync(new Uri(downloadUrl), manifest);
            Status($"Actualizacion descargada: {installerPath}", InfoBarSeverity.Success);

            var result = await new ContentDialog
            {
                XamlRoot = XamlRoot,
                Title = $"Actualizacion {manifest.Version} descargada",
                Content = string.IsNullOrWhiteSpace(manifest.Notes)
                    ? installerPath
                    : $"{manifest.Notes}\n\n{installerPath}",
                PrimaryButtonText = "Instalar ahora",
                SecondaryButtonText = "Abrir carpeta",
                CloseButtonText = "Despues"
            }.ShowAsync();

            if (result == ContentDialogResult.Primary)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = installerPath,
                    UseShellExecute = true
                });
            }
            else if (result == ContentDialogResult.Secondary)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Path.GetDirectoryName(installerPath)!,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            if (manual)
            {
                Status($"No se pudo comprobar la actualizacion: {ex.Message}", InfoBarSeverity.Error);
                await Windows.System.Launcher.LaunchUriAsync(new Uri(DefaultReleaseUrl));
            }
            else
            {
                Status("No se pudo comprobar actualizaciones automaticamente.", InfoBarSeverity.Warning);
            }
        }
    }

    private static async Task<UpdateManifest> FetchUpdateManifestAsync(Uri updateUri)
    {
        using var client = CreateUpdateClient(download: false);
        var json = await client.GetStringAsync(updateUri);
        if (IsGitHubLatestReleaseApi(updateUri))
        {
            var release = JsonSerializer.Deserialize<GitHubRelease>(json, JsonOptions()) ?? new GitHubRelease();
            var installerAsset = release.Assets.FirstOrDefault(asset =>
                asset.Name.Equals("EditorDePlantillasSetup.exe", StringComparison.OrdinalIgnoreCase));

            return new UpdateManifest
            {
                Version = release.TagName.TrimStart('v', 'V'),
                DownloadUrl = installerAsset?.Url ?? installerAsset?.BrowserDownloadUrl ?? "",
                ReleaseUrl = FirstNonEmpty(release.HtmlUrl, DefaultReleaseUrl),
                FileName = installerAsset?.Name ?? "EditorDePlantillasSetup.exe",
                Notes = release.Body
            };
        }

        return JsonSerializer.Deserialize<UpdateManifest>(json, JsonOptions()) ?? new UpdateManifest();
    }

    private static async Task<string> DownloadUpdateAsync(Uri downloadUri, UpdateManifest manifest)
    {
        using var client = CreateUpdateClient(download: IsGitHubAssetApi(downloadUri));
        using var response = await client.GetAsync(downloadUri, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        var fileName = SafeFileName(FirstNonEmpty(
            manifest.FileName,
            Path.GetFileName(downloadUri.LocalPath),
            $"EditorDePlantillasSetup-{manifest.Version}.exe"));
        if (!fileName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            fileName = Path.ChangeExtension(fileName, ".exe");
        }

        var updatesDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EditorDePlantillasUpdates");
        Directory.CreateDirectory(updatesDir);

        var destination = Path.Combine(updatesDir, fileName);
        await using var output = File.Create(destination);
        await response.Content.CopyToAsync(output);
        return destination;
    }

    private async Task ShowReleaseFallbackAsync(UpdateManifest manifest, bool manual)
    {
        var releaseUrl = FirstNonEmpty(manifest.ReleaseUrl, DefaultReleaseUrl);
        if (!manual)
        {
            Status($"Actualizacion {manifest.Version} disponible.", InfoBarSeverity.Informational);
        }

        var result = await new ContentDialog
        {
            XamlRoot = XamlRoot,
            Title = $"Actualizacion disponible {manifest.Version}",
            Content = "No hay descarga automatica configurada para esta version.",
            PrimaryButtonText = "Abrir GitHub",
            CloseButtonText = "Despues"
        }.ShowAsync();

        if (result == ContentDialogResult.Primary)
        {
            await Windows.System.Launcher.LaunchUriAsync(new Uri(releaseUrl));
        }
    }

    private static HttpClient CreateUpdateClient(bool download)
    {
        var client = new HttpClient();
        client.DefaultRequestHeaders.UserAgent.ParseAdd("EditorDePlantillas-Updater/1.0");
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(download ? "application/octet-stream" : "application/json"));

        var token = Environment.GetEnvironmentVariable("EDITOR_PLANTILLAS_GITHUB_TOKEN");
        if (!string.IsNullOrWhiteSpace(token))
        {
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        return client;
    }

    private static bool SupportsAutomaticUpdate(Uri uri)
    {
        return uri.AbsolutePath.EndsWith(".json", StringComparison.OrdinalIgnoreCase)
            || IsGitHubLatestReleaseApi(uri);
    }

    private static bool IsGitHubLatestReleaseApi(Uri uri)
    {
        return uri.Host.Equals("api.github.com", StringComparison.OrdinalIgnoreCase)
            && uri.AbsolutePath.Contains("/releases/latest", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsGitHubAssetApi(Uri uri)
    {
        return uri.Host.Equals("api.github.com", StringComparison.OrdinalIgnoreCase)
            && uri.AbsolutePath.Contains("/releases/assets/", StringComparison.OrdinalIgnoreCase);
    }

    private static string SafeFileName(string fileName)
    {
        foreach (var invalid in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(invalid, '_');
        }

        return string.IsNullOrWhiteSpace(fileName) ? "EditorDePlantillasSetup.exe" : fileName;
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static bool IsNewerVersion(string version, string currentVersion)
    {
        return Version.TryParse(version.TrimStart('v', 'V'), out var latest)
            && Version.TryParse(currentVersion.TrimStart('v', 'V'), out var current)
            && latest > current;
    }

    private static string GetCurrentAppVersion()
    {
        var informationalVersion = typeof(MainPage).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (!string.IsNullOrWhiteSpace(informationalVersion))
        {
            return informationalVersion.Split('+')[0].TrimStart('v', 'V');
        }

        return typeof(MainPage).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";
    }

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e)
    {
        _zoom = Math.Max(0.55, _zoom - 0.1);
        UpdateZoom();
    }

    private void ZoomInButton_Click(object sender, RoutedEventArgs e)
    {
        _zoom = Math.Min(1.7, _zoom + 0.1);
        UpdateZoom();
    }

    private void ResetZoomButton_Click(object sender, RoutedEventArgs e)
    {
        _zoom = 1.0;
        UpdateZoom();
    }

    private void SidebarToggleButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleSidebar();
    }

    private void SidebarToggleKeyboard_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        ToggleSidebar();
        args.Handled = true;
    }

    private void ToggleSidebar()
    {
        _sidebarVisible = !_sidebarVisible;
        UpdateSidebarState(animate: true);
    }

    private void UpdateSidebarState(bool animate = false)
    {
        var targetWidth = _sidebarVisible ? SidebarExpandedWidth : SidebarCollapsedWidth;
        ToolTipService.SetToolTip(
            SidebarToggleButton,
            _sidebarVisible ? "Ocultar barra lateral    Ctrl+B" : "Mostrar barra lateral    Ctrl+B");

        if (!animate)
        {
            StopSidebarAnimation();
            ApplySidebarWidth(targetWidth, completed: true);
            return;
        }

        StartSidebarAnimation(targetWidth);
    }

    private void StartSidebarAnimation(double targetWidth)
    {
        StopSidebarAnimation();

        _sidebarAnimationStart = SidebarColumn.Width.Value;
        _sidebarAnimationTarget = targetWidth;
        _sidebarAnimationRunning = true;
        _sidebarAnimationClock.Restart();

        SidebarPanel.Visibility = Visibility.Visible;
        SidebarPanel.Opacity = targetWidth > _sidebarAnimationStart ? 0 : 1;
        CompositionTarget.Rendering += SidebarAnimation_Rendering;
    }

    private void SidebarAnimation_Rendering(object? sender, object e)
    {
        if (!_sidebarAnimationRunning)
        {
            return;
        }

        var rawProgress = Math.Clamp(_sidebarAnimationClock.Elapsed.TotalMilliseconds / SidebarAnimationDuration.TotalMilliseconds, 0, 1);
        var eased = 1 - Math.Pow(1 - rawProgress, 3);
        var width = _sidebarAnimationStart + (_sidebarAnimationTarget - _sidebarAnimationStart) * eased;
        var showing = _sidebarAnimationTarget > _sidebarAnimationStart;

        ApplySidebarWidth(width, completed: false);
        SidebarPanel.Opacity = showing ? eased : 1 - eased;

        if (rawProgress >= 1)
        {
            StopSidebarAnimation();
            ApplySidebarWidth(_sidebarAnimationTarget, completed: true);
        }
    }

    private void StopSidebarAnimation()
    {
        if (!_sidebarAnimationRunning)
        {
            return;
        }

        CompositionTarget.Rendering -= SidebarAnimation_Rendering;
        _sidebarAnimationClock.Stop();
        _sidebarAnimationRunning = false;
    }

    private void ApplySidebarWidth(double width, bool completed)
    {
        SidebarColumn.Width = new GridLength(Math.Max(SidebarCollapsedWidth, width));
        UpdateSidebarClip(width);

        if (!completed)
        {
            return;
        }

        if (_sidebarVisible)
        {
            SidebarPanel.Visibility = Visibility.Visible;
            SidebarPanel.Opacity = 1;
        }
        else
        {
            SidebarPanel.Visibility = Visibility.Collapsed;
            SidebarPanel.Opacity = 0;
        }
    }

    private void UpdateSidebarClip(double width)
    {
        SidebarPanel.Clip = new RectangleGeometry
        {
            Rect = new Rect(0, 0, Math.Max(0, width), Math.Max(0, SidebarPanel.ActualHeight))
        };
    }

    private void UpdateZoom()
    {
        ZoomLabel.Text = $"{(int)Math.Round(_zoom * 100)}%";
        PreviewPage.RenderTransform = new ScaleTransform
        {
            ScaleX = _zoom,
            ScaleY = _zoom
        };
    }

    private void Status(string message, InfoBarSeverity severity)
    {
        StatusBar.Message = message;
        StatusBar.Severity = severity;
        StatusBar.IsOpen = true;
    }

    private static string ForcedOutputName(TemplateProfile profile)
    {
        var documentName = string.IsNullOrWhiteSpace(profile.Label)
            ? Path.GetFileNameWithoutExtension(profile.SourcePath)
            : profile.Label.Trim();
        return $"{SanitizeFileName(documentName)} - EDITADO.pdf";
    }

    private static string SanitizeFileName(string fileName)
    {
        foreach (var invalid in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(invalid, '_');
        }

        return string.IsNullOrWhiteSpace(fileName) ? "DOCUMENTO" : fileName;
    }

    private static string Shorten(string text, int max)
    {
        var clean = string.Join(" ", text.Split(Array.Empty<char>(), StringSplitOptions.RemoveEmptyEntries));
        return clean.Length <= max ? clean : clean[..Math.Max(0, max - 3)] + "...";
    }

    private static string SanitizeBallValue(string value)
    {
        var digits = new string(value.Where(char.IsDigit).ToArray());
        if (digits.Length == 0)
        {
            return "00";
        }
        return digits.Length > 2 ? digits[^2..] : digits.PadLeft(2, '0');
    }
}

public sealed class AppConfig
{
    [JsonPropertyName("output_dir")]
    public string OutputDir { get; set; } = MainPageDefault.OutputDir;
    [JsonPropertyName("update_url")]
    public string UpdateUrl { get; set; } = MainPageDefault.UpdateUrl;
    [JsonPropertyName("documents")]
    public List<TemplateProfile> Documents { get; set; } = new();
}

internal static class MainPageDefault
{
    public const string OutputDir = @"C:\Users\lauta\Downloads\docs\docs\GUIAS DE CAMBIOS\EXPORTADOS";
    public const string UpdateUrl = "https://api.github.com/repos/lauty23/editor-de-plantillas/releases/latest";
}

public sealed class UpdateManifest
{
    [JsonPropertyName("version")]
    public string Version { get; set; } = "";
    [JsonPropertyName("download_url")]
    public string DownloadUrl { get; set; } = "";
    [JsonPropertyName("release_url")]
    public string ReleaseUrl { get; set; } = "";
    [JsonPropertyName("file_name")]
    public string FileName { get; set; } = "";
    [JsonPropertyName("notes")]
    public string Notes { get; set; } = "";
}

public sealed class GitHubRelease
{
    [JsonPropertyName("tag_name")]
    public string TagName { get; set; } = "";
    [JsonPropertyName("html_url")]
    public string HtmlUrl { get; set; } = "";
    [JsonPropertyName("body")]
    public string Body { get; set; } = "";
    [JsonPropertyName("assets")]
    public List<GitHubAsset> Assets { get; set; } = new();
}

public sealed class GitHubAsset
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";
    [JsonPropertyName("url")]
    public string Url { get; set; } = "";
    [JsonPropertyName("browser_download_url")]
    public string BrowserDownloadUrl { get; set; } = "";
}

public sealed class TemplateProfile
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";
    [JsonPropertyName("label")]
    public string Label { get; set; } = "";
    [JsonPropertyName("kind")]
    public string Kind { get; set; } = "";
    [JsonPropertyName("source_path")]
    public string SourcePath { get; set; } = "";
    [JsonPropertyName("output_name")]
    public string OutputName { get; set; } = "";
}

public sealed class EditableField
{
    public string Label { get; set; } = "";
    public string Original { get; set; } = "";
    public string Value { get; set; } = "";
    public string PartName { get; set; } = "";
    public int GroupIndexInPart { get; set; }
}

public sealed class BallField
{
    public string Label { get; set; } = "";
    public string MediaPath { get; set; } = "";
    public List<string> Values { get; set; } = new() { "16", "17", "26", "38", "41", "05" };
}

public static class RadicadoLogic
{
    public static string Build(IReadOnlyList<EditableField> fields, DateTime timestamp)
    {
        var documentDate = TryReadDocumentDate(fields) ?? timestamp.Date;
        return $"Radicado:000{documentDate:yyyy}{documentDate:dd}{timestamp:HHmmss}-01";
    }

    private static DateTime? TryReadDocumentDate(IReadOnlyList<EditableField> fields)
    {
        foreach (var field in fields)
        {
            var match = Regex.Match(
                field.Value,
                @"(?<day>\d{1,2})\s+de\s+(?<month>[a-zA-ZáéíóúÁÉÍÓÚñÑ]+)\s+de\s+(?<year>\d{4})",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                continue;
            }

            if (!int.TryParse(match.Groups["day"].Value, out var day)
                || !int.TryParse(match.Groups["year"].Value, out var year))
            {
                continue;
            }

            var month = MonthNumber(match.Groups["month"].Value);
            if (month is null)
            {
                continue;
            }

            try
            {
                return new DateTime(year, month.Value, day);
            }
            catch
            {
                return null;
            }
        }

        return null;
    }

    private static int? MonthNumber(string value)
    {
        return NormalizeMonth(value) switch
        {
            "enero" => 1,
            "febrero" => 2,
            "marzo" => 3,
            "abril" => 4,
            "mayo" => 5,
            "junio" => 6,
            "julio" => 7,
            "agosto" => 8,
            "septiembre" => 9,
            "setiembre" => 9,
            "octubre" => 10,
            "noviembre" => 11,
            "diciembre" => 12,
            _ => null
        };
    }

    private static string NormalizeMonth(string value)
    {
        return value.Trim().ToLowerInvariant()
            .Replace('á', 'a')
            .Replace('é', 'e')
            .Replace('í', 'i')
            .Replace('ó', 'o')
            .Replace('ú', 'u');
    }
}

public static class SpanishProofing
{
    private static readonly IReadOnlyDictionary<string, string> Replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["bogota"] = "bogotá",
        ["sabado"] = "sábado",
        ["miercoles"] = "miércoles",
        ["cedula"] = "cédula",
        ["valido"] = "válido",
        ["validos"] = "válidos",
        ["publico"] = "público",
        ["publica"] = "pública",
        ["basico"] = "básico",
        ["basicos"] = "básicos",
        ["reclamacion"] = "reclamación",
        ["informacion"] = "información",
        ["unica"] = "única",
        ["unico"] = "único",
        ["exclusivamente"] = "exclusivamente",
        ["identificacion"] = "identificación",
        ["ciudadania"] = "ciudadanía",
        ["extranjeria"] = "extranjería",
        ["numero"] = "número",
        ["electronico"] = "electrónico",
        ["electronica"] = "electrónica",
        ["electronicamente"] = "electrónicamente",
        ["depositos"] = "depósitos",
        ["debera"] = "deberá",
        ["sera"] = "será",
        ["hara"] = "hará",
        ["credito"] = "crédito",
        ["ministerio"] = "ministerio",
        ["sebastian"] = "sebastián",
        ["alarcon"] = "alarcón"
    };

    public static string Correct(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var corrected = text;
        foreach (var replacement in Replacements)
        {
            corrected = Regex.Replace(
                corrected,
                $@"(?<![\p{{L}}]){Regex.Escape(replacement.Key)}(?![\p{{L}}])",
                match => MatchCase(match.Value, replacement.Value),
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        return corrected;
    }

    private static string MatchCase(string source, string corrected)
    {
        if (source.All(character => !char.IsLetter(character) || char.IsUpper(character)))
        {
            return corrected.ToUpperInvariant();
        }

        if (source.Length > 0 && char.IsUpper(source[0]))
        {
            return char.ToUpperInvariant(corrected[0]) + corrected[1..];
        }

        return corrected;
    }
}

public static class DocxTemplate
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Xml = "http://www.w3.org/XML/1998/namespace";

    public static List<EditableField> ReadFields(string path)
    {
        var fields = new List<EditableField>();
        using var archive = ZipFile.OpenRead(path);
        foreach (var entry in archive.Entries.Where(IsWordXmlPart))
        {
            using var stream = entry.Open();
            var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            var localIndex = 0;
            foreach (var paragraph in document.Descendants(W + "p"))
            {
                foreach (var group in HighlightGroups(paragraph))
                {
                    var value = SpanishProofing.Correct(string.Concat(group.SelectMany(run => run.Descendants(W + "t")).Select(text => text.Value)));
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    fields.Add(new EditableField
                    {
                        Label = $"Campo {fields.Count + 1:00}",
                        Original = value,
                        Value = value,
                        PartName = entry.FullName,
                        GroupIndexInPart = localIndex
                    });
                    localIndex++;
                }
            }
        }
        return fields;
    }

    public static List<BallField> ReadBallFields(string path)
    {
        var fields = new List<BallField>();
        using var archive = ZipFile.OpenRead(path);
        foreach (var entry in archive.Entries.Where(IsImagePart))
        {
            try
            {
                using var stream = entry.Open();
                using var bitmap = new System.Drawing.Bitmap(stream);
                var boxes = BallImage.DetectBallBoxes(bitmap);
                if (boxes.Count == 6)
                {
                    fields.Add(new BallField
                    {
                        Label = $"Numeros de balotas {fields.Count + 1}",
                        MediaPath = entry.FullName
                    });
                }
            }
            catch
            {
                // Some embedded image formats are not handled by GDI+. They are ignored.
            }
        }
        return fields;
    }

    public static void Export(string sourcePath, string destinationPath, IReadOnlyList<EditableField> fields, IReadOnlyList<BallField> ballFields, string? signatureImagePath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        if (File.Exists(destinationPath))
        {
            File.Delete(destinationPath);
        }

        var radicado = RadicadoLogic.Build(fields, DateTime.Now);
        using var source = ZipFile.OpenRead(sourcePath);
        using var destination = ZipFile.Open(destinationPath, ZipArchiveMode.Create);
        foreach (var sourceEntry in source.Entries)
        {
            var targetEntry = destination.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
            if (string.IsNullOrEmpty(sourceEntry.Name))
            {
                continue;
            }

            using var input = sourceEntry.Open();
            using var output = targetEntry.Open();
            var ballField = ballFields.FirstOrDefault(item => item.MediaPath == sourceEntry.FullName);
            if (ballField is not null)
            {
                BallImage.Render(input, output, ballField.Values, sourceEntry.FullName);
            }
            else if (SignatureImage.IsSignatureMedia(sourceEntry.FullName) && !string.IsNullOrWhiteSpace(signatureImagePath) && File.Exists(signatureImagePath))
            {
                SignatureImage.Render(input, output, signatureImagePath, sourceEntry.FullName);
            }
            else if (IsWordXmlPart(sourceEntry) || IsProofingXmlPart(sourceEntry))
            {
                var document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
                if (IsWordXmlPart(sourceEntry))
                {
                    ApplyFields(document, sourceEntry.FullName, fields);
                    ApplyRadicado(document, sourceEntry.FullName, radicado);
                }
                ApplySpanishProofing(document, sourceEntry.FullName);
                if (IsWordXmlPart(sourceEntry))
                {
                    ApplyDocumentTypography(document, sourceEntry.FullName);
                }
                document.Save(output, SaveOptions.DisableFormatting);
            }
            else
            {
                input.CopyTo(output);
            }
        }
    }

    private static void ApplyRadicado(XDocument document, string partName, string radicado)
    {
        if (!partName.Replace('\\', '/').Equals("word/document.xml", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        foreach (var paragraph in document.Descendants(W + "p"))
        {
            var text = string.Concat(paragraph.Descendants(W + "t").Select(item => item.Value)).Trim();
            if (text.StartsWith("Radicado:", StringComparison.OrdinalIgnoreCase))
            {
                ReplaceParagraphText(paragraph, radicado);
            }
        }
    }

    private static void ReplaceParagraphText(XElement paragraph, string value)
    {
        var textElements = paragraph.Descendants(W + "t").ToList();
        if (textElements.Count == 0)
        {
            var run = paragraph.Element(W + "r") ?? new XElement(W + "r");
            if (run.Parent is null)
            {
                paragraph.Add(run);
            }
            run.Add(new XElement(W + "t", value));
            return;
        }

        textElements[0].Value = SpanishProofing.Correct(value);
        textElements[0].SetAttributeValue(Xml + "space", "preserve");
        foreach (var extra in textElements.Skip(1))
        {
            extra.Value = "";
        }
    }

    private static bool IsWordXmlPart(ZipArchiveEntry entry)
    {
        var name = entry.FullName.Replace('\\', '/');
        return name.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase)
               || (name.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
               || (name.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
               || name.Equals("word/footnotes.xml", StringComparison.OrdinalIgnoreCase)
               || name.Equals("word/endnotes.xml", StringComparison.OrdinalIgnoreCase)
               || name.Equals("word/comments.xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsProofingXmlPart(ZipArchiveEntry entry)
    {
        var name = entry.FullName.Replace('\\', '/');
        return name.Equals("word/settings.xml", StringComparison.OrdinalIgnoreCase)
               || name.Equals("word/styles.xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsImagePart(ZipArchiveEntry entry)
    {
        var name = entry.FullName.Replace('\\', '/');
        return name.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase)
               && (name.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                   || name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                   || name.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsHighlighted(XElement run)
    {
        var highlight = run.Element(W + "rPr")?.Element(W + "highlight");
        if (highlight is null)
        {
            return false;
        }

        var value = highlight.Attribute(W + "val")?.Value;
        return string.IsNullOrWhiteSpace(value)
               || value.Equals("yellow", StringComparison.OrdinalIgnoreCase)
               || value.Equals("default", StringComparison.OrdinalIgnoreCase);
    }

    private static void ApplyFields(XDocument document, string partName, IReadOnlyList<EditableField> fields)
    {
        var localIndex = 0;
        foreach (var paragraph in document.Descendants(W + "p"))
        {
            foreach (var group in HighlightGroups(paragraph))
            {
                var field = fields.FirstOrDefault(item => item.PartName == partName && item.GroupIndexInPart == localIndex);
                if (field is not null && group.Count > 0)
                {
                    ReplaceRunText(group[0], field.Value);
                    foreach (var extraRun in group.Skip(1))
                    {
                        foreach (var text in extraRun.Descendants(W + "t").ToList())
                        {
                            text.Value = "";
                        }
                    }
                }
                foreach (var run in group)
                {
                    run.Element(W + "rPr")?.Element(W + "highlight")?.Remove();
                }
                localIndex++;
            }
        }
    }

    private static void ApplyDocumentTypography(XDocument document, string partName)
    {
        var normalizedPart = partName.Replace('\\', '/');
        if (normalizedPart.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var paragraph in document.Descendants(W + "p"))
            {
                var text = string.Concat(paragraph.Descendants(W + "t").Select(item => item.Value)).Trim();
                foreach (var run in paragraph.Elements(W + "r"))
                {
                    if (string.IsNullOrEmpty(string.Concat(run.Descendants(W + "t").Select(item => item.Value))))
                    {
                        continue;
                    }

                    if (text.Contains("COMUNICADO OFICIAL", StringComparison.OrdinalIgnoreCase))
                    {
                        SetRunTypography(run, "Arial", 14, forceBold: true);
                    }
                    else if (text.StartsWith("Radicado:", StringComparison.OrdinalIgnoreCase))
                    {
                        SetRunTypography(run, "Arial", 8, forceBold: true);
                    }
                    else
                    {
                        SetRunTypography(run, "Arial", 11);
                    }
                }
            }
            return;
        }

        if (normalizedPart.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var run in document.Descendants(W + "r"))
            {
                if (!string.IsNullOrEmpty(string.Concat(run.Descendants(W + "t").Select(item => item.Value))))
                {
                    SetRunTypography(run, "Calibri", 8);
                }
            }
        }
    }

    private static void ApplySpanishProofing(XDocument document, string partName)
    {
        foreach (var proofError in document.Descendants(W + "proofErr").ToList())
        {
            proofError.Remove();
        }

        foreach (var text in document.Descendants(W + "t"))
        {
            text.Value = SpanishProofing.Correct(text.Value);
        }

        foreach (var run in document.Descendants(W + "r"))
        {
            SetRunLanguage(run);
        }

        var normalizedPart = partName.Replace('\\', '/');
        if (normalizedPart.Equals("word/settings.xml", StringComparison.OrdinalIgnoreCase))
        {
            ApplySpanishSettings(document);
        }
        else if (normalizedPart.Equals("word/styles.xml", StringComparison.OrdinalIgnoreCase))
        {
            ApplySpanishStyleDefaults(document);
        }
    }

    private static void ApplySpanishSettings(XDocument document)
    {
        var root = document.Root;
        if (root is null)
        {
            return;
        }

        root.Element(W + "proofState")?.Remove();
        var themeFontLang = root.Element(W + "themeFontLang");
        if (themeFontLang is null)
        {
            themeFontLang = new XElement(W + "themeFontLang");
            root.Add(themeFontLang);
        }

        themeFontLang.SetAttributeValue(W + "val", "es-CO");
        themeFontLang.SetAttributeValue(W + "eastAsia", "es-CO");
        themeFontLang.SetAttributeValue(W + "bidi", "es-CO");
    }

    private static void ApplySpanishStyleDefaults(XDocument document)
    {
        var root = document.Root;
        if (root is null)
        {
            return;
        }

        var docDefaults = GetOrAdd(root, "docDefaults");
        var runDefault = GetOrAdd(docDefaults, "rPrDefault");
        var runProperties = GetOrAdd(runDefault, "rPr");
        SetLanguage(runProperties);
    }

    private static void SetRunTypography(XElement run, string family, int sizePoints, bool forceBold = false)
    {
        var runProperties = run.Element(W + "rPr");
        if (runProperties is null)
        {
            runProperties = new XElement(W + "rPr");
            run.AddFirst(runProperties);
        }

        var fonts = runProperties.Element(W + "rFonts");
        if (fonts is null)
        {
            fonts = new XElement(W + "rFonts");
            runProperties.AddFirst(fonts);
        }
        fonts.SetAttributeValue(W + "ascii", family);
        fonts.SetAttributeValue(W + "hAnsi", family);
        fonts.SetAttributeValue(W + "cs", family);

        SetRunPropertyValue(runProperties, "sz", (sizePoints * 2).ToString());
        SetRunPropertyValue(runProperties, "szCs", (sizePoints * 2).ToString());

        if (forceBold && runProperties.Element(W + "b") is null)
        {
            runProperties.Add(new XElement(W + "b"));
        }

        SetLanguage(runProperties);
    }

    private static void SetRunPropertyValue(XElement runProperties, string localName, string value)
    {
        var element = runProperties.Element(W + localName);
        if (element is null)
        {
            element = new XElement(W + localName);
            runProperties.Add(element);
        }
        element.SetAttributeValue(W + "val", value);
    }

    private static void SetRunLanguage(XElement run)
    {
        var runProperties = run.Element(W + "rPr");
        if (runProperties is null)
        {
            runProperties = new XElement(W + "rPr");
            run.AddFirst(runProperties);
        }

        SetLanguage(runProperties);
    }

    private static void SetLanguage(XElement runProperties)
    {
        var language = runProperties.Element(W + "lang");
        if (language is null)
        {
            language = new XElement(W + "lang");
            runProperties.Add(language);
        }

        language.SetAttributeValue(W + "val", "es-CO");
        language.SetAttributeValue(W + "eastAsia", "es-CO");
        language.SetAttributeValue(W + "bidi", "es-CO");
    }

    private static XElement GetOrAdd(XElement parent, string localName)
    {
        var child = parent.Element(W + localName);
        if (child is null)
        {
            child = new XElement(W + localName);
            parent.Add(child);
        }

        return child;
    }

    public static string ParagraphText(XElement paragraph, string partName, IReadOnlyList<EditableField> fields, ref int localIndex)
    {
        var text = "";
        var groups = HighlightGroups(paragraph);
        var groupLookup = groups.SelectMany((group, index) => group.Select(run => (Run: run, Index: index))).ToDictionary(item => item.Run, item => item.Index);
        var emitted = new HashSet<int>();

        foreach (var run in paragraph.Elements(W + "r"))
        {
            if (groupLookup.TryGetValue(run, out var groupIndex))
            {
                if (emitted.Add(groupIndex))
                {
                    EditableField? field = null;
                    foreach (var candidate in fields)
                    {
                        if (candidate.PartName == partName && candidate.GroupIndexInPart == localIndex)
                        {
                            field = candidate;
                            break;
                        }
                    }
                    text += field?.Value ?? string.Concat(groups[groupIndex].SelectMany(item => item.Descendants(W + "t")).Select(item => item.Value));
                    localIndex++;
                }
            }
            else
            {
                text += string.Concat(run.Descendants(W + "t").Select(item => item.Value));
            }
        }

        return text;
    }

    private static List<List<XElement>> HighlightGroups(XElement paragraph)
    {
        var groups = new List<List<XElement>>();
        var current = new List<XElement>();

        void Flush()
        {
            if (current.Count > 0)
            {
                groups.Add(current.ToList());
                current.Clear();
            }
        }

        foreach (var run in paragraph.Elements(W + "r"))
        {
            var text = string.Concat(run.Descendants(W + "t").Select(item => item.Value));
            if (IsHighlighted(run) && !string.IsNullOrEmpty(text))
            {
                current.Add(run);
            }
            else
            {
                Flush();
            }
        }

        Flush();
        return groups;
    }

    private static void ReplaceRunText(XElement run, string value)
    {
        var textElements = run.Descendants(W + "t").ToList();
        if (textElements.Count == 0)
        {
            run.Add(new XElement(W + "t", SpanishProofing.Correct(value)));
            return;
        }

        textElements[0].Value = SpanishProofing.Correct(value);
        textElements[0].SetAttributeValue(Xml + "space", "preserve");
        foreach (var extra in textElements.Skip(1))
        {
            extra.Remove();
        }
    }
}

public static class PdfExport
{
    private const int PagePixelWidth = 794;
    private const int PagePixelHeight = 1123;
    private const double PagePointWidth = 595.28;
    private const double PagePointHeight = 841.89;

    public static void FromDocx(
        string sourcePath,
        string destinationPath,
        IReadOnlyList<EditableField> fields,
        IReadOnlyList<BallField> ballFields,
        string? signatureImagePath)
    {
        var tempImage = Path.Combine(Path.GetTempPath(), $"editor_plantillas_pdf_{Guid.NewGuid():N}.png");
        try
        {
            DocxPreviewRenderer.Render(sourcePath, tempImage, fields, ballFields, signatureImagePath);
            FromImage(tempImage, destinationPath);
        }
        finally
        {
            if (File.Exists(tempImage))
            {
                File.Delete(tempImage);
            }
        }
    }

    public static void FromImage(string imagePath, string destinationPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
        using var source = new System.Drawing.Bitmap(imagePath);
        using var page = new System.Drawing.Bitmap(PagePixelWidth, PagePixelHeight, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
        using (var graphics = System.Drawing.Graphics.FromImage(page))
        {
            graphics.Clear(System.Drawing.Color.White);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var imageRatio = source.Width / (double)source.Height;
            var pageRatio = PagePixelWidth / (double)PagePixelHeight;
            var margin = Math.Abs(imageRatio - pageRatio) < 0.03 ? 0 : 36;
            var available = new System.Drawing.Rectangle(margin, margin, PagePixelWidth - margin * 2, PagePixelHeight - margin * 2);
            var target = FitInside(source.Width, source.Height, available);
            graphics.DrawImage(source, target);
        }

        WriteSingleImagePdf(page, destinationPath);
    }

    private static System.Drawing.Rectangle FitInside(int sourceWidth, int sourceHeight, System.Drawing.Rectangle bounds)
    {
        var scale = Math.Min(bounds.Width / (double)sourceWidth, bounds.Height / (double)sourceHeight);
        var width = Math.Max(1, (int)Math.Round(sourceWidth * scale));
        var height = Math.Max(1, (int)Math.Round(sourceHeight * scale));
        return new System.Drawing.Rectangle(
            bounds.Left + (bounds.Width - width) / 2,
            bounds.Top + (bounds.Height - height) / 2,
            width,
            height);
    }

    private static void WriteSingleImagePdf(System.Drawing.Bitmap page, string destinationPath)
    {
        using var jpegStream = new MemoryStream();
        var encoder = ImageCodecInfo.GetImageEncoders().First(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
        using (var parameters = new EncoderParameters(1))
        {
            parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 94L);
            page.Save(jpegStream, encoder, parameters);
        }

        var imageBytes = jpegStream.ToArray();
        var content = $"q\n{PagePointWidth:0.##} 0 0 {PagePointHeight:0.##} 0 0 cm\n/Im0 Do\nQ\n";
        var contentBytes = System.Text.Encoding.ASCII.GetBytes(content);

        using var output = File.Create(destinationPath);
        var offsets = new List<long> { 0 };
        WriteAscii(output, "%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n");
        WriteObject(output, offsets, 1, "<< /Type /Catalog /Pages 2 0 R >>");
        WriteObject(output, offsets, 2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
        WriteObject(output, offsets, 3, $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PagePointWidth:0.##} {PagePointHeight:0.##}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>");
        WriteImageObject(output, offsets, 4, page.Width, page.Height, imageBytes);
        WriteStreamObject(output, offsets, 5, contentBytes);

        var xrefStart = output.Position;
        WriteAscii(output, $"xref\n0 {offsets.Count}\n");
        WriteAscii(output, "0000000000 65535 f \n");
        foreach (var offset in offsets.Skip(1))
        {
            WriteAscii(output, $"{offset:0000000000} 00000 n \n");
        }
        WriteAscii(output, $"trailer\n<< /Size {offsets.Count} /Root 1 0 R >>\nstartxref\n{xrefStart}\n%%EOF\n");
    }

    private static void WriteObject(Stream output, List<long> offsets, int number, string body)
    {
        offsets.Add(output.Position);
        WriteAscii(output, $"{number} 0 obj\n{body}\nendobj\n");
    }

    private static void WriteImageObject(Stream output, List<long> offsets, int number, int width, int height, byte[] imageBytes)
    {
        offsets.Add(output.Position);
        WriteAscii(output, $"{number} 0 obj\n<< /Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {imageBytes.Length} >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
    }

    private static void WriteStreamObject(Stream output, List<long> offsets, int number, byte[] bytes)
    {
        offsets.Add(output.Position);
        WriteAscii(output, $"{number} 0 obj\n<< /Length {bytes.Length} >>\nstream\n");
        output.Write(bytes, 0, bytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
    }

    private static void WriteAscii(Stream output, string value)
    {
        var bytes = System.Text.Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}

public static class BallImage
{
    private const string YellowBallAsset = "yellow-ball.png";
    private const string RedBallAsset = "red-ball.png";

    public static List<System.Drawing.Rectangle> DetectBallBoxes(System.Drawing.Bitmap image)
    {
        var width = image.Width;
        var height = image.Height;
        var mask = new bool[width * height];
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var color = image.GetPixel(x, y);
                var max = Math.Max(color.R, Math.Max(color.G, color.B));
                var min = Math.Min(color.R, Math.Min(color.G, color.B));
                if (color.A > 20 && !(color.R > 245 && color.G > 245 && color.B > 245) && (max - min > 24 || min < 90))
                {
                    mask[y * width + x] = true;
                }
            }
        }

        var seen = new bool[mask.Length];
        var boxes = new List<(System.Drawing.Rectangle Rect, int Count)>();
        for (var start = 0; start < mask.Length; start++)
        {
            if (!mask[start] || seen[start])
            {
                continue;
            }

            var stack = new Stack<int>();
            stack.Push(start);
            seen[start] = true;
            var minX = width;
            var minY = height;
            var maxX = 0;
            var maxY = 0;
            var count = 0;
            while (stack.Count > 0)
            {
                var point = stack.Pop();
                var x = point % width;
                var y = point / width;
                count++;
                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);

                foreach (var neighbor in new[] { point - 1, point + 1, point - width, point + width })
                {
                    if (neighbor < 0 || neighbor >= mask.Length)
                    {
                        continue;
                    }
                    if ((neighbor == point - 1 && x == 0) || (neighbor == point + 1 && x == width - 1))
                    {
                        continue;
                    }
                    if (mask[neighbor] && !seen[neighbor])
                    {
                        seen[neighbor] = true;
                        stack.Push(neighbor);
                    }
                }
            }

            var rect = System.Drawing.Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
            if (count > 250 && rect.Width > 30 && rect.Height > 30)
            {
                boxes.Add((rect, count));
            }
        }

        return boxes
            .OrderBy(item => item.Rect.Left)
            .Take(6)
            .Select(item => System.Drawing.Rectangle.Inflate(item.Rect, 4, 4))
            .ToList();
    }

    public static void Render(Stream originalStream, Stream output, IReadOnlyList<string> values, string targetName)
    {
        using var original = new System.Drawing.Bitmap(originalStream);
        var boxes = DetectBallBoxes(original);
        if (boxes.Count != 6)
        {
            var diameter = (int)(original.Height * 0.68);
            var gap = Math.Max(8, (original.Width - diameter * 6) / 7);
            var top = (original.Height - diameter) / 2;
            boxes = Enumerable.Range(0, 6)
                .Select(index => new System.Drawing.Rectangle(gap + index * (diameter + gap), top, diameter, diameter))
                .ToList();
        }

        using var image = new System.Drawing.Bitmap(original.Width, original.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = System.Drawing.Graphics.FromImage(image);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.Clear(System.Drawing.Color.Transparent);
        for (var index = 0; index < boxes.Count; index++)
        {
            DrawBall(graphics, boxes[index], Sanitize(values.ElementAtOrDefault(index)), index == boxes.Count - 1);
        }

        if (targetName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || targetName.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
        {
            using var white = new System.Drawing.Bitmap(image.Width, image.Height);
            using var whiteGraphics = System.Drawing.Graphics.FromImage(white);
            whiteGraphics.Clear(System.Drawing.Color.White);
            whiteGraphics.DrawImage(image, 0, 0);
            white.Save(output, ImageFormat.Jpeg);
        }
        else
        {
            image.Save(output, ImageFormat.Png);
        }
    }

    private static void DrawBall(System.Drawing.Graphics graphics, System.Drawing.Rectangle box, string value, bool red)
    {
        var ballSize = Math.Min(70, Math.Min(box.Width, box.Height));
        if (ballSize < 40)
        {
            ballSize = Math.Min(box.Width, box.Height);
        }
        var left = box.Left + (box.Width - ballSize) / 2;
        var top = box.Top + (box.Height - ballSize) / 2;
        var rect = new System.Drawing.Rectangle(left, top, ballSize, ballSize);
        var scale = ballSize / 70f;

        using (var shadow = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(126, 0, 0, 0)))
        {
            graphics.FillEllipse(shadow, rect.Left + 5 * scale, rect.Top + 5 * scale, rect.Width, rect.Height);
        }

        using var sprite = LoadBallSprite(red);
        if (sprite is not null)
        {
            graphics.DrawImage(sprite, rect);
        }
        else
        {
            DrawFallbackBall(graphics, rect, red, scale);
        }

        using var font = new System.Drawing.Font(
            "Segoe UI",
            Math.Max(18, 32 * scale),
            System.Drawing.FontStyle.Bold,
            System.Drawing.GraphicsUnit.Pixel);
        using var textBrush = new System.Drawing.SolidBrush(red ? System.Drawing.Color.White : System.Drawing.Color.Black);
        using var format = new System.Drawing.StringFormat
        {
            Alignment = System.Drawing.StringAlignment.Center,
            LineAlignment = System.Drawing.StringAlignment.Near
        };
        var textRect = new System.Drawing.RectangleF(rect.Left, rect.Top + 8 * scale, rect.Width, rect.Height - 8 * scale);
        graphics.DrawString(value, font, textBrush, textRect, format);

        var measured = graphics.MeasureString(value, font);
        var underlineWidth = Math.Min(rect.Width * 0.48f, measured.Width * 0.82f);
        var underlineY = rect.Top + 48 * scale;
        var underlineStart = rect.Left + (rect.Width - underlineWidth) / 2f;
        using var underlinePen = new System.Drawing.Pen(red ? System.Drawing.Color.White : System.Drawing.Color.Black, Math.Max(2f, 2.25f * scale));
        underlinePen.StartCap = LineCap.Round;
        underlinePen.EndCap = LineCap.Round;
        graphics.DrawLine(underlinePen, underlineStart, underlineY, underlineStart + underlineWidth, underlineY);
    }

    public static void DrawPreviewBall(System.Drawing.Graphics graphics, System.Drawing.Rectangle box, string value, bool red)
    {
        DrawBall(graphics, box, Sanitize(value), red);
    }

    private static System.Drawing.Bitmap? LoadBallSprite(bool red)
    {
        var fileName = red ? RedBallAsset : YellowBallAsset;
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "Assets", fileName),
            Path.Combine(AppContext.BaseDirectory, fileName),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", fileName)
        };

        foreach (var candidate in candidates.Select(Path.GetFullPath))
        {
            if (!File.Exists(candidate))
            {
                continue;
            }

            using var stream = File.OpenRead(candidate);
            using var loaded = new System.Drawing.Bitmap(stream);
            return new System.Drawing.Bitmap(loaded);
        }

        return null;
    }

    private static void DrawFallbackBall(System.Drawing.Graphics graphics, System.Drawing.Rectangle rect, bool red, float scale)
    {
        using var path = new GraphicsPath();
        path.AddEllipse(rect);
        using var brush = new PathGradientBrush(path)
        {
            CenterPoint = new System.Drawing.PointF(rect.Left + rect.Width * 0.36f, rect.Top + rect.Height * 0.26f),
            CenterColor = red ? System.Drawing.Color.FromArgb(255, 255, 74, 62) : System.Drawing.Color.FromArgb(255, 255, 231, 88),
            SurroundColors = new[] { red ? System.Drawing.Color.FromArgb(255, 138, 0, 0) : System.Drawing.Color.FromArgb(255, 209, 125, 0) }
        };
        graphics.FillPath(brush, path);

        using var lowerShade = new System.Drawing.Drawing2D.LinearGradientBrush(
            rect,
            System.Drawing.Color.FromArgb(0, 255, 255, 255),
            red ? System.Drawing.Color.FromArgb(120, 75, 0, 0) : System.Drawing.Color.FromArgb(110, 161, 79, 0),
            LinearGradientMode.Vertical);
        graphics.FillEllipse(lowerShade, rect);

        using var rim = new System.Drawing.Pen(System.Drawing.Color.FromArgb(75, 255, 255, 255), Math.Max(1f, scale));
        graphics.DrawEllipse(rim, rect);
    }

    private static string Sanitize(string? value)
    {
        var digits = new string((value ?? "").Where(char.IsDigit).ToArray());
        if (digits.Length == 0)
        {
            return "00";
        }
        return digits.Length > 2 ? digits[^2..] : digits.PadLeft(2, '0');
    }
}

public static class SignatureImage
{
    public static bool IsSignatureMedia(string mediaPath)
    {
        return mediaPath.Replace('\\', '/').EndsWith("/image2.png", StringComparison.OrdinalIgnoreCase);
    }

    public static void Render(Stream originalStream, Stream output, string signaturePath, string targetName)
    {
        using var original = new System.Drawing.Bitmap(originalStream);
        using var replacement = LoadCropped(signaturePath);
        var originalBounds = DetectContentBounds(original, 0) ?? new System.Drawing.Rectangle(
            original.Width / 3,
            original.Height / 3,
            original.Width / 3,
            original.Height / 3);

        using var canvas = new System.Drawing.Bitmap(original.Width, original.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = System.Drawing.Graphics.FromImage(canvas);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.Clear(System.Drawing.Color.Transparent);

        var target = FitInside(replacement.Width, replacement.Height, originalBounds);
        graphics.DrawImage(replacement, target);

        if (targetName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || targetName.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
        {
            using var white = new System.Drawing.Bitmap(canvas.Width, canvas.Height);
            using var whiteGraphics = System.Drawing.Graphics.FromImage(white);
            whiteGraphics.Clear(System.Drawing.Color.White);
            whiteGraphics.DrawImage(canvas, 0, 0);
            white.Save(output, ImageFormat.Jpeg);
        }
        else
        {
            canvas.Save(output, ImageFormat.Png);
        }
    }

    public static System.Drawing.Bitmap LoadCropped(string signaturePath)
    {
        using var loaded = new System.Drawing.Bitmap(signaturePath);
        var bounds = DetectContentBounds(loaded, 12);
        if (bounds is null)
        {
            return new System.Drawing.Bitmap(loaded);
        }

        return loaded.Clone(bounds.Value, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
    }

    private static System.Drawing.Rectangle FitInside(int sourceWidth, int sourceHeight, System.Drawing.Rectangle target)
    {
        var scale = Math.Min(target.Width / (double)sourceWidth, target.Height / (double)sourceHeight);
        var width = Math.Max(1, (int)Math.Round(sourceWidth * scale));
        var height = Math.Max(1, (int)Math.Round(sourceHeight * scale));
        return new System.Drawing.Rectangle(
            target.Left + (target.Width - width) / 2,
            target.Top + (target.Height - height) / 2,
            width,
            height);
    }

    private static System.Drawing.Rectangle? DetectContentBounds(System.Drawing.Bitmap source, int margin)
    {
        var left = source.Width;
        var top = source.Height;
        var right = 0;
        var bottom = 0;

        for (var y = 0; y < source.Height; y++)
        {
            for (var x = 0; x < source.Width; x++)
            {
                var color = source.GetPixel(x, y);
                if (color.A > 12 && color.R < 190 && color.G < 190 && color.B < 190)
                {
                    left = Math.Min(left, x);
                    top = Math.Min(top, y);
                    right = Math.Max(right, x);
                    bottom = Math.Max(bottom, y);
                }
            }
        }

        if (right <= left || bottom <= top)
        {
            return null;
        }

        left = Math.Max(0, left - margin);
        top = Math.Max(0, top - margin);
        right = Math.Min(source.Width - 1, right + margin);
        bottom = Math.Min(source.Height - 1, bottom + margin);
        return System.Drawing.Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
    }
}

public static class DocxPreviewRenderer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const int PageWidth = 794;
    private const int PageHeight = 1123;
    private const int MarginLeft = 108;
    private const int MarginRight = 70;
    private const int HeaderLeft = 72;
    private const int HeaderRight = 58;

    public static void Render(string docxPath, string outputPath, IReadOnlyList<EditableField> fields, IReadOnlyList<BallField> ballFields, string? signatureImagePath)
    {
        using var archive = ZipFile.OpenRead(docxPath);
        using var page = new System.Drawing.Bitmap(PageWidth, PageHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = System.Drawing.Graphics.FromImage(page);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        graphics.Clear(System.Drawing.Color.White);

        var relationships = LoadRelationships(archive);
        var radicado = RadicadoLogic.Build(fields, DateTime.Now);
        RenderHeader(archive, graphics, relationships, ballFields);
        RenderBody(archive, graphics, relationships, fields, ballFields, radicado, signatureImagePath);
        RenderFooter(archive, graphics, relationships, fields, ballFields);

        page.Save(outputPath, ImageFormat.Png);
    }

    private static Dictionary<string, Dictionary<string, string>> LoadRelationships(ZipArchive archive)
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var relEntry in archive.Entries.Where(entry => entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                using var stream = relEntry.Open();
                var document = XDocument.Load(stream);
                var partName = PartNameFromRels(relEntry.FullName);
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var relationship in document.Descendants(Rel + "Relationship"))
                {
                    var id = relationship.Attribute("Id")?.Value;
                    var target = relationship.Attribute("Target")?.Value;
                    if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                    {
                        map[id] = ResolveTarget(partName, target);
                    }
                }
                result[partName] = map;
            }
            catch
            {
                // Ignore malformed relationship parts in preview mode.
            }
        }
        return result;
    }

    private static string PartNameFromRels(string relPath)
    {
        relPath = relPath.Replace('\\', '/');
        var marker = "/_rels/";
        var markerIndex = relPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            return relPath;
        }
        var directory = relPath[..markerIndex];
        var file = relPath[(markerIndex + marker.Length)..];
        if (file.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
        {
            file = file[..^5];
        }
        return $"{directory}/{file}".TrimStart('/');
    }

    private static string ResolveTarget(string partName, string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith('/'))
        {
            return target.TrimStart('/');
        }

        var slash = partName.LastIndexOf('/');
        var baseDir = slash >= 0 ? partName[..slash] : "";
        var pieces = $"{baseDir}/{target}".Split('/', StringSplitOptions.RemoveEmptyEntries);
        var stack = new Stack<string>();
        foreach (var piece in pieces)
        {
            if (piece == ".")
            {
                continue;
            }
            if (piece == "..")
            {
                if (stack.Count > 0)
                {
                    stack.Pop();
                }
                continue;
            }
            stack.Push(piece);
        }
        return string.Join("/", stack.Reverse());
    }

    private static void RenderHeader(
        ZipArchive archive,
        System.Drawing.Graphics graphics,
        Dictionary<string, Dictionary<string, string>> relationships,
        IReadOnlyList<BallField> ballFields)
    {
        var headerParts = archive.Entries
            .Where(entry => entry.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName)
            .ToList();

        var images = new List<string>();
        foreach (var header in headerParts)
        {
            foreach (var mediaPath in DrawingTargets(header, relationships))
            {
                if (!images.Contains(mediaPath, StringComparer.OrdinalIgnoreCase))
                {
                    images.Add(mediaPath);
                }
            }
        }

        var balotoLogo = images.FirstOrDefault(item => item.EndsWith("image4.png", StringComparison.OrdinalIgnoreCase));
        var nuevoPaisLogo = images.FirstOrDefault(item => item.EndsWith("image3.png", StringComparison.OrdinalIgnoreCase));

        if (balotoLogo is not null)
        {
            using var image = LoadPreviewImage(archive, balotoLogo, ballFields);
            if (image is not null)
            {
                var w = 150;
                var h = (int)(image.Height * (w / (double)image.Width));
                graphics.DrawImage(image, HeaderLeft, 34, w, h);
            }
        }

        if (nuevoPaisLogo is not null)
        {
            using var image = LoadPreviewImage(archive, nuevoPaisLogo, ballFields);
            if (image is not null)
            {
                var cropX = (int)(image.Width * 0.57);
                var crop = new System.Drawing.Rectangle(cropX, 0, image.Width - cropX, image.Height);
                var w = 128;
                var h = (int)(crop.Height * (w / (double)crop.Width));
                graphics.DrawImage(image, new System.Drawing.Rectangle(PageWidth - HeaderRight - w, 34, w, h), crop, System.Drawing.GraphicsUnit.Pixel);
            }
        }

        if (balotoLogo is not null || nuevoPaisLogo is not null)
        {
            return;
        }

        for (var index = 0; index < images.Count; index++)
        {
            var image = LoadPreviewImage(archive, images[index], ballFields);
            if (image is null)
            {
                continue;
            }
            using (image)
            {
                var maxW = index == 0 ? 155 : 135;
                var maxH = 58;
                var scale = Math.Min(maxW / (double)image.Width, maxH / (double)image.Height);
                var w = (int)(image.Width * scale);
                var h = (int)(image.Height * scale);
                var x = index == 0 ? HeaderLeft : PageWidth - HeaderRight - w;
                var y = 38 + index / 2 * 64;
                graphics.DrawImage(image, x, y, w, h);
            }
        }
    }

    private static void RenderBody(
        ZipArchive archive,
        System.Drawing.Graphics graphics,
        Dictionary<string, Dictionary<string, string>> relationships,
        IReadOnlyList<EditableField> fields,
        IReadOnlyList<BallField> ballFields,
        string radicado,
        string? signatureImagePath)
    {
        var entry = archive.GetEntry("word/document.xml");
        if (entry is null)
        {
            return;
        }

        using var stream = entry.Open();
        var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        var y = 124f;
        var partFieldIndex = 0;
        var reachedIntro = false;
        var blankParagraphsToSkip = 0;
        foreach (var paragraph in document.Descendants(W + "p"))
        {
            var drawings = DrawingTargets(paragraph, "word/document.xml", relationships).ToList();
            var segments = ParagraphSegments(paragraph, "word/document.xml", fields, ref partFieldIndex);
            var visibleText = string.Concat(segments.Select(segment => segment.Text)).Trim();
            if (visibleText.StartsWith("Radicado:", StringComparison.OrdinalIgnoreCase))
            {
                segments = ReplacePreviewParagraphText(segments, radicado);
                visibleText = radicado;
            }
            segments = ApplyPreviewTypography(segments, "word/document.xml", visibleText);

            if (string.IsNullOrWhiteSpace(visibleText) && drawings.Count == 0)
            {
                if (blankParagraphsToSkip > 0)
                {
                    blankParagraphsToSkip--;
                    continue;
                }

                y += reachedIntro ? 4 : 14;
                continue;
            }

            if (!string.IsNullOrWhiteSpace(visibleText))
            {
            var alignment = ParagraphAlignment(paragraph);
            var isNumberedParagraph = paragraph.Element(W + "pPr")?.Element(W + "numPr") is not null;
            var left = MarginLeft + TwipsToPixels(IntAttribute(paragraph.Element(W + "pPr")?.Element(W + "ind"), "left") ?? 0);
            if (isNumberedParagraph)
            {
                left = MarginLeft + 24;
            }
            var width = PageWidth - left - MarginRight;
                var consumed = DrawStyledParagraph(graphics, segments, left, y, width, alignment);
                y += consumed + ParagraphSpacingAfter(paragraph);

                if (visibleText.StartsWith("Baloto se permite", StringComparison.OrdinalIgnoreCase))
                {
                    reachedIntro = true;
                }
            }

            if (drawings.Count > 0)
            {
                y += string.IsNullOrWhiteSpace(visibleText) ? 0 : 14;
                var drewSignature = false;
                foreach (var mediaPath in drawings)
                {
                    var previewBallField = ballFields.FirstOrDefault(item => item.MediaPath.Equals(mediaPath, StringComparison.OrdinalIgnoreCase));
                    if (previewBallField is not null)
                    {
                        y += 18;
                        DrawPreviewBallStrip(graphics, previewBallField.Values, y);
                        y += 87;
                        continue;
                    }

                    var image = LoadPreviewImage(archive, mediaPath, ballFields, signatureImagePath);
                    if (image is null)
                    {
                        continue;
                    }
                    using (image)
                    {
                        var isSignature = mediaPath.EndsWith("image2.png", StringComparison.OrdinalIgnoreCase);
                        var maxW = isSignature ? 178 : 250;
                        var maxH = isSignature ? 92 : 135;
                        var scale = Math.Min(maxW / (double)image.Width, maxH / (double)image.Height);
                        scale = Math.Min(scale, 1.25);
                        var w = (int)(image.Width * scale);
                        var h = (int)(image.Height * scale);
                        var x = MarginLeft;
                        graphics.DrawImage(image, x, y, w, h);
                        y += h + (isSignature ? 8 : 12);
                        drewSignature = drewSignature || isSignature;
                    }
                }

                if (drewSignature)
                {
                    blankParagraphsToSkip = 4;
                }
            }

            if (y > PageHeight - 140)
            {
                break;
            }
        }
    }

    private static void RenderFooter(
        ZipArchive archive,
        System.Drawing.Graphics graphics,
        Dictionary<string, Dictionary<string, string>> relationships,
        IReadOnlyList<EditableField> fields,
        IReadOnlyList<BallField> ballFields)
    {
        var footerParts = archive.Entries
            .Where(entry => entry.FullName.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName)
            .ToList();
        if (footerParts.Count == 0)
        {
            return;
        }

        var y = PageHeight - 112f;
        foreach (var footer in footerParts.Take(1))
        {
            using var stream = footer.Open();
            var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            var partIndex = 0;
            foreach (var paragraph in document.Descendants(W + "p"))
            {
                var segments = ParagraphSegments(paragraph, footer.FullName, fields, ref partIndex);
                segments = ApplyPreviewTypography(segments, footer.FullName, string.Concat(segments.Select(segment => segment.Text)).Trim());
                if (!string.IsNullOrWhiteSpace(string.Concat(segments.Select(segment => segment.Text))))
                {
                    y += DrawStyledParagraph(graphics, segments, MarginLeft, y, PageWidth - MarginLeft - MarginRight - 110, PreviewAlignment.Left);
                }
                foreach (var mediaPath in DrawingTargets(paragraph, footer.FullName, relationships))
                {
                    var image = LoadPreviewImage(archive, mediaPath, ballFields);
                    if (image is null)
                    {
                        continue;
                    }
                    using (image)
                    {
                        var scale = Math.Min(105 / (double)image.Width, 52 / (double)image.Height);
                        var w = (int)(image.Width * scale);
                        var h = (int)(image.Height * scale);
                        graphics.DrawImage(image, PageWidth - MarginRight - w, PageHeight - 76, w, h);
                    }
                }
            }
        }
    }

    private static IEnumerable<string> DrawingTargets(ZipArchiveEntry part, Dictionary<string, Dictionary<string, string>> relationships)
    {
        using var stream = part.Open();
        var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        foreach (var paragraph in document.Descendants(W + "p"))
        {
            foreach (var target in DrawingTargets(paragraph, part.FullName, relationships))
            {
                yield return target;
            }
        }
    }

    private static IEnumerable<string> DrawingTargets(XElement paragraph, string partName, Dictionary<string, Dictionary<string, string>> relationships)
    {
        if (!relationships.TryGetValue(partName, out var partRelationships))
        {
            yield break;
        }

        foreach (var blip in paragraph.Descendants().Where(element => element.Name.LocalName == "blip"))
        {
            var relId = blip.Attribute(R + "embed")?.Value;
            if (!string.IsNullOrWhiteSpace(relId) && partRelationships.TryGetValue(relId, out var target))
            {
                yield return target;
            }
        }
    }

    private static System.Drawing.Image? LoadPreviewImage(ZipArchive archive, string mediaPath, IReadOnlyList<BallField> ballFields, string? signatureImagePath = null)
    {
        var entry = archive.GetEntry(mediaPath);
        if (entry is null)
        {
            return null;
        }

        if (SignatureImage.IsSignatureMedia(mediaPath) && !string.IsNullOrWhiteSpace(signatureImagePath) && File.Exists(signatureImagePath))
        {
            return SignatureImage.LoadCropped(signatureImagePath);
        }

        var ballField = ballFields.FirstOrDefault(item => item.MediaPath.Equals(mediaPath, StringComparison.OrdinalIgnoreCase));
        if (ballField is not null)
        {
            using var input = entry.Open();
            var output = new MemoryStream();
            BallImage.Render(input, output, ballField.Values, mediaPath);
            output.Position = 0;
            return System.Drawing.Image.FromStream(output);
        }

        using var stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        memory.Position = 0;
        using var loaded = new System.Drawing.Bitmap(memory);
        if (mediaPath.EndsWith("image2.png", StringComparison.OrdinalIgnoreCase))
        {
            return CropNonWhiteContent(loaded, 18);
        }

        return new System.Drawing.Bitmap(loaded);
    }

    private static void DrawPreviewBallStrip(System.Drawing.Graphics graphics, IReadOnlyList<string> values, float y)
    {
        const int ballSize = 70;
        const int gap = 31;
        var totalWidth = ballSize * 6 + gap * 5;
        var x = (PageWidth - totalWidth) / 2;

        for (var index = 0; index < 6; index++)
        {
            var rect = new System.Drawing.Rectangle(x + index * (ballSize + gap), (int)y, ballSize, ballSize);
            BallImage.DrawPreviewBall(graphics, rect, values.ElementAtOrDefault(index) ?? "00", index == 5);
        }
    }

    private static System.Drawing.Bitmap CropNonWhiteContent(System.Drawing.Bitmap source, int margin)
    {
        var left = source.Width;
        var top = source.Height;
        var right = 0;
        var bottom = 0;

        for (var y = 0; y < source.Height; y++)
        {
            for (var x = 0; x < source.Width; x++)
            {
                var color = source.GetPixel(x, y);
                if (color.A > 12 && color.R < 190 && color.G < 190 && color.B < 190)
                {
                    left = Math.Min(left, x);
                    top = Math.Min(top, y);
                    right = Math.Max(right, x);
                    bottom = Math.Max(bottom, y);
                }
            }
        }

        if (right <= left || bottom <= top)
        {
            return new System.Drawing.Bitmap(source);
        }

        left = Math.Max(0, left - margin);
        top = Math.Max(0, top - margin);
        right = Math.Min(source.Width - 1, right + margin);
        bottom = Math.Min(source.Height - 1, bottom + margin);
        var crop = System.Drawing.Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
        return source.Clone(crop, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
    }

    private static List<PreviewTextRun> ParagraphSegments(XElement paragraph, string partName, IReadOnlyList<EditableField> fields, ref int localIndex)
    {
        var segments = new List<PreviewTextRun>();
        var groups = HighlightGroups(paragraph);
        var groupLookup = groups
            .SelectMany((group, index) => group.Select(run => (Run: run, Index: index)))
            .ToDictionary(item => item.Run, item => item.Index);
        var emitted = new HashSet<int>();

        foreach (var run in paragraph.Elements(W + "r"))
        {
            if (groupLookup.TryGetValue(run, out var groupIndex))
            {
                if (!emitted.Add(groupIndex))
                {
                    continue;
                }

                EditableField? field = null;
                foreach (var candidate in fields)
                {
                    if (candidate.PartName == partName && candidate.GroupIndexInPart == localIndex)
                    {
                        field = candidate;
                        break;
                    }
                }
                var value = SpanishProofing.Correct(field?.Value ?? string.Concat(groups[groupIndex].Select(RunText)));
                if (!string.IsNullOrEmpty(value))
                {
                    segments.Add(new PreviewTextRun(value, ReadRunFormat(groups[groupIndex][0], paragraph)));
                }
                localIndex++;
            }
            else
            {
                var value = SpanishProofing.Correct(RunText(run));
                if (!string.IsNullOrEmpty(value))
                {
                    segments.Add(new PreviewTextRun(value, ReadRunFormat(run, paragraph)));
                }
            }
        }

        if (paragraph.Element(W + "pPr")?.Element(W + "numPr") is not null
            && segments.Count > 0
            && !string.Concat(segments.Select(segment => segment.Text)).TrimStart().StartsWith("•", StringComparison.Ordinal))
        {
            segments.Insert(0, new PreviewTextRun("•\t", segments[0].Format));
        }

        return segments;
    }

    private static List<PreviewTextRun> ReplacePreviewParagraphText(IReadOnlyList<PreviewTextRun> segments, string value)
    {
        var format = segments.FirstOrDefault()?.Format ?? new PreviewRunFormat("Arial", 8, true, false, false);
        return new List<PreviewTextRun> { new(SpanishProofing.Correct(value), format) };
    }

    private static List<PreviewTextRun> ApplyPreviewTypography(IReadOnlyList<PreviewTextRun> segments, string partName, string paragraphText)
    {
        var normalizedPart = partName.Replace('\\', '/');
        var isFooter = normalizedPart.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase);
        var isTitle = paragraphText.Contains("COMUNICADO OFICIAL", StringComparison.OrdinalIgnoreCase);
        var isRadicado = paragraphText.StartsWith("Radicado:", StringComparison.OrdinalIgnoreCase);

        return segments.Select(segment =>
        {
            var current = segment.Format;
            PreviewRunFormat format;
            if (isFooter)
            {
                format = current with { FontFamily = "Calibri", SizePoints = 8 };
            }
            else if (isTitle)
            {
                format = current with { FontFamily = "Arial", SizePoints = 14, Bold = true };
            }
            else if (isRadicado)
            {
                format = current with { FontFamily = "Arial", SizePoints = 8, Bold = true };
            }
            else
            {
                format = current with { FontFamily = "Arial", SizePoints = 11 };
            }

            return segment with { Format = format };
        }).ToList();
    }

    private static List<List<XElement>> HighlightGroups(XElement paragraph)
    {
        var groups = new List<List<XElement>>();
        var current = new List<XElement>();

        void Flush()
        {
            if (current.Count > 0)
            {
                groups.Add(current.ToList());
                current.Clear();
            }
        }

        foreach (var run in paragraph.Elements(W + "r"))
        {
            var text = RunText(run);
            if (IsHighlighted(run) && !string.IsNullOrEmpty(text))
            {
                current.Add(run);
            }
            else
            {
                Flush();
            }
        }

        Flush();
        return groups;
    }

    private static bool IsHighlighted(XElement run)
    {
        var highlight = run.Element(W + "rPr")?.Element(W + "highlight");
        if (highlight is null)
        {
            return false;
        }

        var value = highlight.Attribute(W + "val")?.Value;
        return string.IsNullOrWhiteSpace(value)
               || value.Equals("yellow", StringComparison.OrdinalIgnoreCase)
               || value.Equals("default", StringComparison.OrdinalIgnoreCase);
    }

    private static string RunText(XElement run)
    {
        var pieces = new List<string>();
        foreach (var child in run.Elements())
        {
            if (child.Name == W + "t")
            {
                pieces.Add(child.Value);
            }
            else if (child.Name == W + "tab")
            {
                pieces.Add("\t");
            }
            else if (child.Name == W + "br")
            {
                pieces.Add("\n");
            }
        }
        return string.Concat(pieces);
    }

    private static PreviewRunFormat ReadRunFormat(XElement run, XElement paragraph)
    {
        var paragraphRunProperties = paragraph.Element(W + "pPr")?.Element(W + "rPr");
        var runProperties = run.Element(W + "rPr");

        var family = FontFamily(runProperties) ?? FontFamily(paragraphRunProperties) ?? "Calibri";
        var size = HalfPointSize(runProperties) ?? HalfPointSize(paragraphRunProperties) ?? 11f;
        var bold = OnOff(paragraphRunProperties?.Element(W + "b")) ?? false;
        bold = OnOff(runProperties?.Element(W + "b")) ?? bold;
        var italic = OnOff(paragraphRunProperties?.Element(W + "i")) ?? false;
        italic = OnOff(runProperties?.Element(W + "i")) ?? italic;
        var underline = UnderlineOn(paragraphRunProperties?.Element(W + "u")) ?? false;
        underline = UnderlineOn(runProperties?.Element(W + "u")) ?? underline;

        return new PreviewRunFormat(family, size, bold, italic, underline);
    }

    private static string? FontFamily(XElement? runProperties)
    {
        var fonts = runProperties?.Element(W + "rFonts");
        return fonts?.Attribute(W + "ascii")?.Value
               ?? fonts?.Attribute(W + "hAnsi")?.Value
               ?? fonts?.Attribute(W + "cs")?.Value;
    }

    private static float? HalfPointSize(XElement? runProperties)
    {
        var value = runProperties?.Element(W + "sz")?.Attribute(W + "val")?.Value;
        return int.TryParse(value, out var halfPoints) ? Math.Max(1f, halfPoints / 2f) : null;
    }

    private static bool? OnOff(XElement? element)
    {
        if (element is null)
        {
            return null;
        }

        var value = element.Attribute(W + "val")?.Value;
        return value is "0" or "false" or "off" ? false : true;
    }

    private static bool? UnderlineOn(XElement? element)
    {
        if (element is null)
        {
            return null;
        }

        var value = element.Attribute(W + "val")?.Value;
        return value is "none" or "0" or "false" ? false : true;
    }

    private static PreviewAlignment ParagraphAlignment(XElement paragraph)
    {
        var value = paragraph.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val")?.Value;
        return value switch
        {
            "center" => PreviewAlignment.Center,
            "right" => PreviewAlignment.Right,
            _ => PreviewAlignment.Left
        };
    }

    private static float ParagraphSpacingAfter(XElement paragraph)
    {
        var after = IntAttribute(paragraph.Element(W + "pPr")?.Element(W + "spacing"), "after");
        return after.HasValue ? Math.Min(18f, TwipsToPixels(after.Value)) : 6f;
    }

    private static int? IntAttribute(XElement? element, string name)
    {
        var value = element?.Attribute(W + name)?.Value;
        return int.TryParse(value, out var parsed) ? parsed : null;
    }

    private static float TwipsToPixels(int twips) => twips * 96f / 1440f;

    private static float DrawStyledParagraph(System.Drawing.Graphics graphics, IReadOnlyList<PreviewTextRun> segments, float x, float y, float width, PreviewAlignment alignment)
    {
        var tokens = segments.SelectMany(Tokenize).ToList();
        if (tokens.Count == 0)
        {
            return 0;
        }

        var lines = new List<List<PreviewTextToken>>();
        var currentLine = new List<PreviewTextToken>();
        var currentWidth = 0f;

        foreach (var token in tokens)
        {
            if (token.Text.Contains('\n'))
            {
                FlushLine();
                continue;
            }

            if (string.IsNullOrWhiteSpace(token.Text) && currentLine.Count == 0)
            {
                continue;
            }

            var measured = MeasureToken(graphics, token);
            if (currentLine.Count > 0 && currentWidth + measured > width)
            {
                FlushLine();
                if (string.IsNullOrWhiteSpace(token.Text))
                {
                    continue;
                }
            }

            currentLine.Add(token);
            currentWidth += measured;
        }

        FlushLine();

        var startY = y;
        foreach (var line in lines)
        {
            var lineWidth = line.Sum(token => MeasureToken(graphics, token));
            var lineHeight = Math.Max(1f, line.Max(token => FontHeight(graphics, token.Run.Format)) * 1.05f);
            var drawX = alignment switch
            {
                PreviewAlignment.Center => x + Math.Max(0, (width - lineWidth) / 2f),
                PreviewAlignment.Right => x + Math.Max(0, width - lineWidth),
                _ => x
            };

            foreach (var token in line)
            {
                using var font = CreateFont(token.Run.Format);
                var tokenWidth = MeasureToken(graphics, token);
                if (!string.IsNullOrWhiteSpace(token.Text) || token.Run.Format.Underline)
                {
                    using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
                    using var format = TypographicFormat();
                    graphics.DrawString(DrawableTokenText(token), font, brush, drawX, y, format);
                }
                drawX += tokenWidth;
            }

            y += lineHeight;
        }

        return Math.Max(0, y - startY);

        void FlushLine()
        {
            while (currentLine.Count > 0 && string.IsNullOrWhiteSpace(currentLine[^1].Text))
            {
                currentLine.RemoveAt(currentLine.Count - 1);
            }

            if (currentLine.Count > 0)
            {
                lines.Add(currentLine.ToList());
                currentLine.Clear();
                currentWidth = 0;
            }
        }
    }

    private static IEnumerable<PreviewTextToken> Tokenize(PreviewTextRun run)
    {
        var normalized = Regex.Replace(run.Text.Replace("\t", " "), @"[^\S\r\n]+", " ");
        foreach (Match match in Regex.Matches(normalized, @"\n|[^\S\r\n]+|\S+"))
        {
            yield return new PreviewTextToken(match.Value, run);
        }
    }

    private static float MeasureToken(System.Drawing.Graphics graphics, PreviewTextToken token)
    {
        using var font = CreateFont(token.Run.Format);
        if (string.IsNullOrWhiteSpace(token.Text) && !token.Run.Format.Underline)
        {
            var spaceCount = Math.Max(1, token.Text.Count(char.IsWhiteSpace));
            return spaceCount * font.SizeInPoints * 96f / 72f * 0.32f;
        }

        using var format = TypographicFormat();
        return graphics.MeasureString(DrawableTokenText(token), font, int.MaxValue, format).Width;
    }

    private static string DrawableTokenText(PreviewTextToken token)
    {
        if (string.IsNullOrWhiteSpace(token.Text) && token.Run.Format.Underline)
        {
            return new string('\u00A0', Math.Max(1, token.Text.Count(char.IsWhiteSpace)));
        }

        return token.Text;
    }

    private static float FontHeight(System.Drawing.Graphics graphics, PreviewRunFormat format)
    {
        using var font = CreateFont(format);
        return font.GetHeight(graphics);
    }

    private static System.Drawing.Font CreateFont(PreviewRunFormat format)
    {
        var style = System.Drawing.FontStyle.Regular;
        if (format.Bold)
        {
            style |= System.Drawing.FontStyle.Bold;
        }
        if (format.Italic)
        {
            style |= System.Drawing.FontStyle.Italic;
        }
        if (format.Underline)
        {
            style |= System.Drawing.FontStyle.Underline;
        }
        try
        {
            return new System.Drawing.Font(format.FontFamily, format.SizePoints, style, System.Drawing.GraphicsUnit.Point);
        }
        catch
        {
            return new System.Drawing.Font("Calibri", format.SizePoints, style, System.Drawing.GraphicsUnit.Point);
        }
    }

    private static System.Drawing.StringFormat TypographicFormat()
    {
        var format = (System.Drawing.StringFormat)System.Drawing.StringFormat.GenericTypographic.Clone();
        format.FormatFlags |= System.Drawing.StringFormatFlags.MeasureTrailingSpaces;
        return format;
    }

    private enum PreviewAlignment
    {
        Left,
        Center,
        Right
    }

    private sealed record PreviewRunFormat(string FontFamily, float SizePoints, bool Bold, bool Italic, bool Underline);
    private sealed record PreviewTextRun(string Text, PreviewRunFormat Format);
    private sealed record PreviewTextToken(string Text, PreviewTextRun Run);
}
