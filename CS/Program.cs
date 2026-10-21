// See https://aka.ms/new-console-template for more information
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using DevExpress.Docs;
using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;

class Program {
    static void Main() {
        while (true) {
            // Show menu for user action selection
            Console.WriteLine("Select an action:");
            Console.WriteLine("1 - Edit color scheme only");
            Console.WriteLine("2 - Edit font scheme only");
            Console.WriteLine("3 - Edit format scheme only");
            Console.WriteLine("4 - Edit shape style");
            Console.WriteLine("5 - Apply ThemeOverride to the first slide");
            Console.WriteLine("6 - Remap master and slide colors");
            Console.WriteLine("7 - Apply custom theme to the entire presentation");
            Console.WriteLine("8 - Customize theme for a specific layout");
            Console.WriteLine("0 - Exit");
            Console.Write("Your choice: ");
            var choice = Console.ReadLine()?.Trim(); ;
            if (choice == "0") break;

            string fileSuffix = choice switch {
                "1" => "colorscheme",
                "2" => "fontscheme",
                "3" => "formatscheme",
                "4" => "shapestyle",
                "5" => "themeoverride",
                "6" => "mapping",
                "7" => "customtheme",
                "8" => "layouttheme",
                _ => "result"
            };
            string pptxPath = Path.Combine(Environment.CurrentDirectory, $"presentation_{fileSuffix}.pptx");
            string pdfPath = Path.Combine(Environment.CurrentDirectory, $"presentation_{fileSuffix}.pdf");

            // Open the presentation file
            using (FileStream fs = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.pptx"))) {
                Presentation presentation = new Presentation(fs);
                switch (choice) {
                    case "1":
                        CustomizeColorScheme(presentation); break;
                    case "2":
                        CustomizeFontScheme(presentation); break;
                    case "3":
                        CustomizeFormatScheme(presentation); break;
                    case "4":
                        EditShapeStyle(presentation); break;
                    case "5":
                        ApplyThemeOverride(presentation); break;
                    case "6":
                        RemapColors(presentation); break;
                    case "7":
                        ApplyCustomTheme(presentation); break;
                    case "8":
                        CustomizeLayoutTheme(presentation); break;
                    default:
                        Console.WriteLine("Unknown choice.");
                        continue;
                }
                // Always export PDF
                ExportToPdfAndOpen(presentation, pdfPath);
                Console.Write($"PDF will be saved as '{pdfPath}'. Do you also want to save the presentation file (pptx)? (y/n): ");
                var saveChoice = Console.ReadLine();
                if (saveChoice?.Trim().ToLower() == "y") {
                    SavePresentationFile(presentation, pptxPath);
                }
            }
            Console.WriteLine();
        }
        Console.WriteLine("Application exited.");
    }

    static void CustomizeColorScheme(Presentation presentation) {
        // Get first slide and its master
        var slide = presentation.Slides[0];
        var mySlideMaster = presentation.SlideMasters.Where(sm => sm.Layouts.Contains(slide.Layout)).FirstOrDefault();
        var myTheme = mySlideMaster?.Theme;
        if (myTheme == null) return;

        // Edit color scheme
        myTheme.ColorScheme.Accent1 = new OfficeColor(Color.Green);
        myTheme.ColorScheme.Accent2 = new OfficeColor(Color.DarkKhaki);
        myTheme.ColorScheme.Accent3 = new OfficeColor(Color.Olive);
        myTheme.ColorScheme.Dark2 = new OfficeColor(Color.Red);
        Console.WriteLine("Color scheme customized for the presentation.");
    }

    static void CustomizeFontScheme(Presentation presentation) {
        // Get first slide and its master
        var slide = presentation.Slides[0];
        var mySlideMaster = presentation.SlideMasters.Where(sm => sm.Layouts.Contains(slide.Layout)).FirstOrDefault();
        var myTheme = mySlideMaster?.Theme;
        if (myTheme == null) return;

        // Edit font scheme
        myTheme.FontScheme.MinorFont = new ThemeFont("Calibri", "Calibri Light", "Calibri");
        myTheme.FontScheme.MajorFont = new ThemeFont("Tahoma");
        Console.WriteLine("Font scheme customized for the presentation.");
    }

    static void CustomizeFormatScheme(Presentation presentation) {

        // Get first slide and its master
        var slide = presentation.Slides[0];
        var mySlideMaster = presentation.SlideMasters.Where(sm => sm.Layouts.Contains(slide.Layout)).FirstOrDefault();
        var myTheme = mySlideMaster?.Theme;
        if (myTheme == null) return;

        // Edit format scheme fills
        myTheme.FormatScheme.Fills.ResetTo(new List<DevExpress.Docs.Office.Fill> {
            new DevExpress.Docs.Office.SolidFill(Color.Red),
            new DevExpress.Docs.Office.SolidFill(Color.Magenta),
            new DevExpress.Docs.Office.SolidFill(Color.Lime)
        });

        // Edit format scheme background fills
        myTheme.FormatScheme.BackgroundFills.ResetTo(new List<DevExpress.Docs.Office.Fill> {
            new DevExpress.Docs.Office.SolidFill(Color.Purple),
            new DevExpress.Docs.Office.SolidFill(Color.Yellow),
            new DevExpress.Docs.Office.SolidFill(Color.DarkGoldenrod)
        });

        // Edit format scheme line styles
        myTheme.FormatScheme.LineStyles.ResetTo(new List<LineStyle> {
            new LineStyle() {
                Width = 50,
                DashType = DevExpress.Docs.Office.LineDashType.Dot,
                Fill = new DevExpress.Docs.Office.SolidFill(Color.Red)
            },
            new LineStyle() {
                Width = 20,
                DashType = DevExpress.Docs.Office.LineDashType.DashDot,
                Fill = new DevExpress.Docs.Office.SolidFill(Color.Magenta)
            },
            new LineStyle() {
                Width = 5,
                DashType = DevExpress.Docs.Office.LineDashType.Solid,
                Fill = new DevExpress.Docs.Office.SolidFill(Color.Lime)
            }
        });

        // Edit format scheme effects
        myTheme.FormatScheme.Effects.ResetTo(new List<ShapeEffectProperties> {
            new ShapeEffectProperties() { Blur = new DevExpress.Docs.Office.BlurEffect(100, true) },
            new ShapeEffectProperties() { InnerShadow = new DevExpress.Docs.Office.InnerShadowEffect() },
            new ShapeEffectProperties() { Glow = new DevExpress.Docs.Office.GlowEffect(15) }
        });
        Console.WriteLine("Format scheme customized for the presentation.");
    }

    static void EditShapeStyle(Presentation presentation) {
        // Edit style of the first shape on the first slide
        var slide = presentation.Slides[0];
        if (slide.Shapes.Count == 0) {
            // Add a rectangle shape if none exists
            var shape = new Shape(ShapeType.Rectangle) {
                X = 100,
                Y = 100,
                Width = 400,
                Height = 200,
                Style = new ShapeStyle() {
                    FillThemeIndex = 1,
                    OutlineThemeIndex = 1,
                },
                TextArea = new DevExpress.Docs.Office.TextArea("Sample themed shape") {
                    Level1ParagraphProperties = new DevExpress.Docs.Office.TextParagraphProperties() {
                        TextProperties = new DevExpress.Docs.Office.TextProperties() {
                            FontSize = 32,
                        },
                    },
                },
            };
            slide.Shapes.Add(shape);
        }
        var targetShape = slide.Shapes[0] as Shape;
        if (targetShape != null) {
            targetShape.Style.FillThemeIndex = 2; // Use theme fill index
            targetShape.Style.OutlineThemeIndex = 3; // Use theme outline index
            targetShape.Style.FontStyle = FontSchemeStyle.Major;
            targetShape.Style.FontColor = new OfficeColor(SchemeColors.Accent3);

            // You can also set direct fill/outline
            // targetShape.Fill = new SolidFill(Color.LightGreen);
            // targetShape.Outline = new Outline() { Width = 5, Fill = new SolidFill(Color.DarkBlue) };
            Console.WriteLine("Shape style edited on the first slide.");
        }
    }

    static void ApplyThemeOverride(Presentation presentation) {
        // Get first slide
        Slide slide = presentation.Slides[0];

        // Create theme override for the slide
        ThemeOverride themeOverride = new ThemeOverride();
        ThemeColorScheme overrideColors = new ThemeColorScheme();
        overrideColors.Accent1 = new OfficeColor(Color.Red);
        overrideColors.Accent2 = new OfficeColor(Color.Red);
        themeOverride.ColorScheme = overrideColors;
        ThemeFormatScheme formatScheme = new ThemeFormatScheme();
        themeOverride.FormatScheme = formatScheme;

        // Set required fills
        themeOverride.FormatScheme.Fills.ResetTo(new List<DevExpress.Docs.Office.Fill> {
            new DevExpress.Docs.Office.SolidFill(Color.Red),
            new DevExpress.Docs.Office.SolidFill(Color.Magenta),
            new DevExpress.Docs.Office.SolidFill(Color.Lime)
        });

        // Apply override to slide
        slide.ThemeOverride = themeOverride;
        Console.WriteLine("ThemeOverride applied to the first slide.");
    }

    static void RemapColors(Presentation presentation) {
        // Remap master color scheme
        SlideMaster slideMaster = presentation.SlideMasters[0];
        ColorMap masterColorMap = slideMaster.ColorMap;
        masterColorMap.Accent1 = ColorSchemeIndexType.Accent3;
        masterColorMap.Accent2 = ColorSchemeIndexType.Accent5;
        masterColorMap.Background1 = ColorSchemeIndexType.Dark2;
        masterColorMap.Text1 = ColorSchemeIndexType.Light2;
        masterColorMap.Hyperlink = ColorSchemeIndexType.Accent6;

        // Remap color scheme for second slide
        Slide slide = presentation.Slides[1];
        ColorMap slideColorMap = new ColorMap();
        slideColorMap.Accent1 = ColorSchemeIndexType.Accent6;
        slideColorMap.Hyperlink = ColorSchemeIndexType.Accent5;
        slide.ColorMapOverride = slideColorMap;
        Console.WriteLine("Master and second slide colors remapped.");
    }

    static void ApplyCustomTheme(Presentation presentation) {
        // Get first slide and its master
        var slide = presentation.Slides[0];
        var mySlideMaster = presentation.SlideMasters.Where(sm => sm.Layouts.Contains(slide.Layout)).FirstOrDefault();
        var myTheme = mySlideMaster?.Theme;
        if (myTheme == null) return;
        // Create custom color scheme
        ThemeColorScheme customColors = new ThemeColorScheme();
        customColors.Name = "Corporate Blue";
        customColors.Dark1 = new OfficeColor(Color.FromArgb(31, 56, 100));
        customColors.Light1 = new OfficeColor(Color.White);
        customColors.Dark2 = new OfficeColor(Color.FromArgb(68, 84, 106));
        customColors.Light2 = new OfficeColor(Color.FromArgb(242, 246, 252));
        customColors.Accent1 = new OfficeColor(Color.FromArgb(68, 114, 196));
        customColors.Accent2 = new OfficeColor(Color.FromArgb(112, 173, 71));
        customColors.Accent3 = new OfficeColor(Color.FromArgb(255, 192, 0));
        customColors.Accent4 = new OfficeColor(Color.FromArgb(158, 72, 14));
        customColors.Accent5 = new OfficeColor(Color.FromArgb(99, 99, 99));
        customColors.Accent6 = new OfficeColor(Color.FromArgb(37, 94, 145));
        customColors.Hyperlink = new OfficeColor(Color.FromArgb(5, 99, 193));
        customColors.FollowedHyperlink = new OfficeColor(Color.FromArgb(149, 79, 114));

        // Create custom font scheme
        ThemeFontScheme customFonts = new ThemeFontScheme();
        customFonts.Name = "Corporate Fonts";
        customFonts.MajorFont = new ThemeFont("Segoe UI", "Arial", "Arial");
        customFonts.MinorFont = new ThemeFont("Calibri", "Times New Roman", "Times New Roman");

        // Create custom format scheme
        ThemeFormatScheme customFormats = new ThemeFormatScheme();

        // Add fills (minimum 3 required)
        customFormats.Fills.Add(new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.FromArgb(68, 114, 196))));
        customFormats.Fills.Add(new DevExpress.Docs.Office.GradientFill(
            DevExpress.Docs.Office.GradientType.Linear, [
                new GradientStop(new OfficeColor(Color.White), 0.0),
                new GradientStop(new OfficeColor(Color.FromArgb(68, 114, 196)), 1.0)
            ]
        ));
        customFormats.Fills.Add(new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.FromArgb(112, 173, 71))));

        // Add background fills (minimum 3 required)
        customFormats.BackgroundFills.Add(new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.White)));
        customFormats.BackgroundFills.Add(new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.FromArgb(242, 246, 252))));
        customFormats.BackgroundFills.Add(new DevExpress.Docs.Office.GradientFill(
            DevExpress.Docs.Office.GradientType.Linear, [
                new GradientStop(new OfficeColor(Color.White), 0.0),
                new GradientStop(new OfficeColor(Color.FromArgb(242, 246, 252)), 1.0)
            ]
        ));

        // Add line styles (minimum 3 required)
        customFormats.LineStyles.Add(new LineStyle() { Width = 1, Fill = new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.FromArgb(68, 114, 196))) });
        customFormats.LineStyles.Add(new LineStyle() { Width = 2, Fill = new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.FromArgb(112, 173, 71))) });
        customFormats.LineStyles.Add(new LineStyle() { Width = 3, Fill = new DevExpress.Docs.Office.SolidFill(new OfficeColor(Color.FromArgb(255, 192, 0))) });

        // Add effects (minimum 3 required)
        customFormats.Effects.Add(new ShapeEffectProperties() { Blur = new DevExpress.Docs.Office.BlurEffect(4, true) });
        customFormats.Effects.Add(new ShapeEffectProperties() { InnerShadow = new DevExpress.Docs.Office.InnerShadowEffect() });
        customFormats.Effects.Add(new ShapeEffectProperties() { Glow = new DevExpress.Docs.Office.GlowEffect(5) });

        // Assign custom theme to master
        myTheme.ColorScheme = customColors;
        myTheme.FontScheme = customFonts;
        myTheme.FormatScheme = customFormats;
        Console.WriteLine("Custom theme applied to the entire presentation.");
    }

    static void CustomizeLayoutTheme(Presentation presentation) {
        // Find target layout by name
        string targetLayoutName = "Title Slide Layout"; // Change to your layout name
        var targetLayout = presentation.SlideMasters
            .SelectMany(sm => sm.Layouts)
            .FirstOrDefault(l => l.Name == targetLayoutName);
        if (targetLayout != null) {

            // Create theme override for layout
            ThemeOverride customTheme = new ThemeOverride();
            ThemeColorScheme customColors = new ThemeColorScheme();
            customColors.Accent1 = new OfficeColor(Color.Purple);
            customColors.Accent2 = new OfficeColor(Color.Orange);
            customTheme.ColorScheme = customColors;

            ThemeFontScheme overrideFonts = new ThemeFontScheme();
            overrideFonts.MajorFont = new ThemeFont("Times New Roman", "Comic Sans MS", "Comic Sans MS");
            overrideFonts.MinorFont = new ThemeFont("Comic Sans MS", "Comic Sans MS");
            customTheme.FontScheme = overrideFonts;

            ThemeFormatScheme customFormats = new ThemeFormatScheme();
            customFormats.Fills.ResetTo(new List<DevExpress.Docs.Office.Fill> {
                new DevExpress.Docs.Office.SolidFill(Color.Purple),
                new DevExpress.Docs.Office.SolidFill(Color.Orange),
                new DevExpress.Docs.Office.SolidFill(Color.Yellow)
            });
            customTheme.FormatScheme = customFormats;

            // Apply theme override to all slides with the target layout
            foreach (var slide in presentation.Slides.Where(s => s.Layout == targetLayout)) {
                slide.ThemeOverride = customTheme;
            }
            Console.WriteLine($"Custom theme applied to all slides with layout '{targetLayoutName}'.");
        } else {
            Console.WriteLine($"Layout '{targetLayoutName}' not found.");
        }
    }

    static void ExportToPdfAndOpen(Presentation presentation, string pdfPath) {
        using (FileStream pdfStream = new FileStream(pdfPath, FileMode.Create)) {
            presentation.ExportToPdf(pdfStream);
        }
        Console.WriteLine($"PDF exported: {pdfPath}");
        // Open the PDF file in the default application
        Process.Start(new ProcessStartInfo {
            FileName = pdfPath,
            UseShellExecute = true
        });
    }

    static void SavePresentationFile(Presentation presentation, string pptxPath) {
        using (FileStream outputStream = new FileStream(pptxPath, FileMode.Create)) {
            presentation.SaveDocument(outputStream);
        }
        Console.WriteLine($"Presentation saved: {pptxPath}");
    }
}