Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports DevExpress.Docs
Imports DevExpress.Docs.Presentation
Imports DevExpress.Docs.Office

Public Module Program
    Public Sub Main()
        While True
            ' Show menu for user action selection
            Console.WriteLine("Select an action:")
            Console.WriteLine("1 - Edit color scheme only")
            Console.WriteLine("2 - Edit font scheme only")
            Console.WriteLine("3 - Edit format scheme only")
            Console.WriteLine("4 - Edit shape style")
            Console.WriteLine("5 - Apply ThemeOverride to the first slide")
            Console.WriteLine("6 - Remap master and slide colors")
            Console.WriteLine("7 - Apply custom theme to the entire presentation")
            Console.WriteLine("8 - Customize theme for a specific layout")
            Console.WriteLine("0 - Exit")
            Console.Write("Your choice: ")
            Dim choice = Console.ReadLine()?.Trim()
            If choice = "0" Then Exit While

            Dim fileSuffix As String
            Select Case choice
                Case "1"
                    fileSuffix = "colorscheme"
                Case "2"
                    fileSuffix = "fontscheme"
                Case "3"
                    fileSuffix = "formatscheme"
                Case "4"
                    fileSuffix = "shapestyle"
                Case "5"
                    fileSuffix = "themeoverride"
                Case "6"
                    fileSuffix = "mapping"
                Case "7"
                    fileSuffix = "customtheme"
                Case "8"
                    fileSuffix = "layouttheme"
                Case Else
                    fileSuffix = "result"
            End Select
            Dim pptxPath As String = Path.Combine(Environment.CurrentDirectory, $"presentation_{fileSuffix}.pptx")
            Dim pdfPath As String = Path.Combine(Environment.CurrentDirectory, $"presentation_{fileSuffix}.pdf")

            ' Open the presentation file
            Using fs As FileStream = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.pptx"))
                Dim presentation As New Presentation(fs)
                Select Case choice
                    Case "1"
                        CustomizeColorScheme(presentation)
                    Case "2"
                        CustomizeFontScheme(presentation)
                    Case "3"
                        CustomizeFormatScheme(presentation)
                    Case "4"
                        EditShapeStyle(presentation)
                    Case "5"
                        ApplyThemeOverride(presentation)
                    Case "6"
                        RemapColors(presentation)
                    Case "7"
                        ApplyCustomTheme(presentation)
                    Case "8"
                        CustomizeLayoutTheme(presentation)
                    Case Else
                        Console.WriteLine("Unknown choice.")
                        Continue While
                End Select
                ' Always export PDF
                ExportToPdfAndOpen(presentation, pdfPath)
                Console.Write($"PDF will be saved as '{pdfPath}'. Do you also want to save the presentation file (pptx)? (y/n): ")
                Dim saveChoice = Console.ReadLine()
                If saveChoice?.Trim().ToLower() = "y" Then
                    SavePresentationFile(presentation, pptxPath)
                End If
            End Using
            Console.WriteLine()
        End While
        Console.WriteLine("Application exited.")
    End Sub

    Private Sub CustomizeColorScheme(presentation As Presentation)
        ' Get first slide and its master
        Dim slide = presentation.Slides(0)
        Dim mySlideMaster = presentation.SlideMasters.Where(Function(sm) sm.Layouts.Contains(slide.Layout)).FirstOrDefault()
        Dim myTheme = mySlideMaster?.Theme
        If myTheme Is Nothing Then Return

        ' Edit color scheme
        myTheme.ColorScheme.Accent1 = New OfficeColor(Color.Green)
        myTheme.ColorScheme.Accent2 = New OfficeColor(Color.DarkKhaki)
        myTheme.ColorScheme.Accent3 = New OfficeColor(Color.Olive)
        myTheme.ColorScheme.Dark2 = New OfficeColor(Color.Red)
        Console.WriteLine("Color scheme customized for the presentation.")
    End Sub

    Private Sub CustomizeFontScheme(presentation As Presentation)
        ' Get first slide and its master
        Dim slide = presentation.Slides(0)
        Dim mySlideMaster = presentation.SlideMasters.Where(Function(sm) sm.Layouts.Contains(slide.Layout)).FirstOrDefault()
        Dim myTheme = mySlideMaster?.Theme
        If myTheme Is Nothing Then Return

        ' Edit font scheme
        myTheme.FontScheme.MinorFont = New ThemeFont("Calibri", "Calibri Light", "Calibri")
        myTheme.FontScheme.MajorFont = New ThemeFont("Tahoma")
        Console.WriteLine("Font scheme customized for the presentation.")
    End Sub

    Private Sub CustomizeFormatScheme(presentation As Presentation)

        ' Get first slide and its master
        Dim slide = presentation.Slides(0)
        Dim mySlideMaster = presentation.SlideMasters.Where(Function(sm) sm.Layouts.Contains(slide.Layout)).FirstOrDefault()
        Dim myTheme = mySlideMaster?.Theme
        If myTheme Is Nothing Then Return

        ' Edit format scheme fills
        myTheme.FormatScheme.Fills.ResetTo(New List(Of DevExpress.Docs.Office.Fill) From {
            New DevExpress.Docs.Office.SolidFill(Color.Red),
            New DevExpress.Docs.Office.SolidFill(Color.Magenta),
            New DevExpress.Docs.Office.SolidFill(Color.Lime)
        })

        ' Edit format scheme background fills
        myTheme.FormatScheme.BackgroundFills.ResetTo(New List(Of DevExpress.Docs.Office.Fill) From {
            New DevExpress.Docs.Office.SolidFill(Color.Purple),
            New DevExpress.Docs.Office.SolidFill(Color.Yellow),
            New DevExpress.Docs.Office.SolidFill(Color.DarkGoldenrod)
        })

        ' Edit format scheme line styles
        myTheme.FormatScheme.LineStyles.ResetTo(New List(Of LineStyle) From {
            New LineStyle() With {
                .Width = 50,
                .DashType = DevExpress.Docs.Office.LineDashType.Dot,
                .Fill = New DevExpress.Docs.Office.SolidFill(Color.Red)
            },
            New LineStyle() With {
                .Width = 20,
                .DashType = DevExpress.Docs.Office.LineDashType.DashDot,
                .Fill = New DevExpress.Docs.Office.SolidFill(Color.Magenta)
            },
            New LineStyle() With {
                .Width = 5,
                .DashType = DevExpress.Docs.Office.LineDashType.Solid,
                .Fill = New DevExpress.Docs.Office.SolidFill(Color.Lime)
            }
        })

        ' Edit format scheme effects
        myTheme.FormatScheme.Effects.ResetTo(New List(Of ShapeEffectProperties) From {
            New ShapeEffectProperties() With {.Blur = New DevExpress.Docs.Office.BlurEffect(100, True)},
            New ShapeEffectProperties() With {.InnerShadow = New DevExpress.Docs.Office.InnerShadowEffect()},
            New ShapeEffectProperties() With {.Glow = New DevExpress.Docs.Office.GlowEffect(15)}
        })
        Console.WriteLine("Format scheme customized for the presentation.")
    End Sub

    Private Sub EditShapeStyle(presentation As Presentation)
        ' Edit style of the first shape on the first slide
        Dim slide = presentation.Slides(0)
        If slide.Shapes.Count = 0 Then
            ' Add a rectangle shape if none exists
            Dim shape As New Shape(ShapeType.Rectangle) With {
                .X = 100,
                .Y = 100,
                .Width = 400,
                .Height = 200,
                .Style = New ShapeStyle() With {
                    .FillThemeIndex = 1,
                    .OutlineThemeIndex = 1
                },
                .TextArea = New DevExpress.Docs.Office.TextArea("Sample themed shape") With {
                    .Level1ParagraphProperties = New DevExpress.Docs.Office.TextParagraphProperties() With {
                        .TextProperties = New DevExpress.Docs.Office.TextProperties() With {
                            .FontSize = 32
                        }
                    }
                }
            }
            slide.Shapes.Add(shape)
        End If
        Dim targetShape = TryCast(slide.Shapes(0), Shape)
        If targetShape IsNot Nothing Then
            targetShape.Style.FillThemeIndex = 2 ' Use theme fill index
            targetShape.Style.OutlineThemeIndex = 3 ' Use theme outline index
            targetShape.Style.FontStyle = FontSchemeStyle.Major
            targetShape.Style.FontColor = New OfficeColor(SchemeColors.Accent3)

            ' You can also set direct fill/outline
            ' targetShape.Fill = New SolidFill(Color.LightGreen)
            ' targetShape.Outline = New Outline() With {.Width = 5, .Fill = New SolidFill(Color.DarkBlue)}
            Console.WriteLine("Shape style edited on the first slide.")
        End If
    End Sub

    Private Sub ApplyThemeOverride(presentation As Presentation)
        ' Get first slide
        Dim slide As Slide = presentation.Slides(0)

        ' Create theme override for the slide
        Dim themeOverride As New ThemeOverride()
        Dim overrideColors As New ThemeColorScheme()
        overrideColors.Accent1 = New OfficeColor(Color.Red)
        overrideColors.Accent2 = New OfficeColor(Color.Red)
        themeOverride.ColorScheme = overrideColors
        Dim formatScheme As New ThemeFormatScheme()
        themeOverride.FormatScheme = formatScheme

        ' Set required fills
        themeOverride.FormatScheme.Fills.ResetTo(New List(Of DevExpress.Docs.Office.Fill) From {
            New DevExpress.Docs.Office.SolidFill(Color.Red),
            New DevExpress.Docs.Office.SolidFill(Color.Magenta),
            New DevExpress.Docs.Office.SolidFill(Color.Lime)
        })

        ' Apply override to slide
        slide.ThemeOverride = themeOverride
        Console.WriteLine("ThemeOverride applied to the first slide.")
    End Sub

    Private Sub RemapColors(presentation As Presentation)
        ' Remap master color scheme
        Dim slideMaster As SlideMaster = presentation.SlideMasters(0)
        Dim masterColorMap As ColorMap = slideMaster.ColorMap
        masterColorMap.Accent1 = ColorSchemeIndexType.Accent3
        masterColorMap.Accent2 = ColorSchemeIndexType.Accent5
        masterColorMap.Background1 = ColorSchemeIndexType.Dark2
        masterColorMap.Text1 = ColorSchemeIndexType.Light2
        masterColorMap.Hyperlink = ColorSchemeIndexType.Accent6

        ' Remap color scheme for second slide
        Dim slide As Slide = presentation.Slides(1)
        Dim slideColorMap As New ColorMap()
        slideColorMap.Accent1 = ColorSchemeIndexType.Accent6
        slideColorMap.Hyperlink = ColorSchemeIndexType.Accent5
        slide.ColorMapOverride = slideColorMap
        Console.WriteLine("Master and second slide colors remapped.")
    End Sub

    Private Sub ApplyCustomTheme(presentation As Presentation)
        ' Get first slide and its master
        Dim slide = presentation.Slides(0)
        Dim mySlideMaster = presentation.SlideMasters.Where(Function(sm) sm.Layouts.Contains(slide.Layout)).FirstOrDefault()
        Dim myTheme = mySlideMaster?.Theme
        If myTheme Is Nothing Then Return
        ' Create custom color scheme
        Dim customColors As New ThemeColorScheme()
        customColors.Name = "Corporate Blue"
        customColors.Dark1 = New OfficeColor(Color.FromArgb(31, 56, 100))
        customColors.Light1 = New OfficeColor(Color.White)
        customColors.Dark2 = New OfficeColor(Color.FromArgb(68, 84, 106))
        customColors.Light2 = New OfficeColor(Color.FromArgb(242, 246, 252))
        customColors.Accent1 = New OfficeColor(Color.FromArgb(68, 114, 196))
        customColors.Accent2 = New OfficeColor(Color.FromArgb(112, 173, 71))
        customColors.Accent3 = New OfficeColor(Color.FromArgb(255, 192, 0))
        customColors.Accent4 = New OfficeColor(Color.FromArgb(158, 72, 14))
        customColors.Accent5 = New OfficeColor(Color.FromArgb(99, 99, 99))
        customColors.Accent6 = New OfficeColor(Color.FromArgb(37, 94, 145))
        customColors.Hyperlink = New OfficeColor(Color.FromArgb(5, 99, 193))
        customColors.FollowedHyperlink = New OfficeColor(Color.FromArgb(149, 79, 114))

        ' Create custom font scheme
        Dim customFonts As New ThemeFontScheme()
        customFonts.Name = "Corporate Fonts"
        customFonts.MajorFont = New ThemeFont("Segoe UI", "Arial", "Arial")
        customFonts.MinorFont = New ThemeFont("Calibri", "Times New Roman", "Times New Roman")

        ' Create custom format scheme
        Dim customFormats As New ThemeFormatScheme()

        ' Add fills (minimum 3 required)
        customFormats.Fills.Add(New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.FromArgb(68, 114, 196))))
        customFormats.Fills.Add(New DevExpress.Docs.Office.GradientFill(
            DevExpress.Docs.Office.GradientType.Linear, New GradientStop() {
                New GradientStop(Color.White, 0),
                New GradientStop(Color.FromArgb(68, 114, 196), 1)
            }
        ))
        customFormats.Fills.Add(New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.FromArgb(112, 173, 71))))

        ' Add background fills (minimum 3 required)
        customFormats.BackgroundFills.Add(New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.White)))
        customFormats.BackgroundFills.Add(New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.FromArgb(242, 246, 252))))
        customFormats.BackgroundFills.Add(New DevExpress.Docs.Office.GradientFill(
            DevExpress.Docs.Office.GradientType.Linear, New GradientStop() {
                New GradientStop(Color.White, 0),
                New GradientStop(Color.FromArgb(242, 246, 252), 1)
            }
        ))

        ' Add line styles (minimum 3 required)
        customFormats.LineStyles.Add(New LineStyle() With {.Width = 1, .Fill = New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.FromArgb(68, 114, 196)))})
        customFormats.LineStyles.Add(New LineStyle() With {.Width = 2, .Fill = New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.FromArgb(112, 173, 71)))})
        customFormats.LineStyles.Add(New LineStyle() With {.Width = 3, .Fill = New DevExpress.Docs.Office.SolidFill(New OfficeColor(Color.FromArgb(255, 192, 0)))})

        ' Add effects (minimum 3 required)
        customFormats.Effects.Add(New ShapeEffectProperties() With {.Blur = New DevExpress.Docs.Office.BlurEffect(4, True)})
        customFormats.Effects.Add(New ShapeEffectProperties() With {.InnerShadow = New DevExpress.Docs.Office.InnerShadowEffect()})
        customFormats.Effects.Add(New ShapeEffectProperties() With {.Glow = New DevExpress.Docs.Office.GlowEffect(5)})

        ' Assign custom theme to master
        myTheme.ColorScheme = customColors
        myTheme.FontScheme = customFonts
        myTheme.FormatScheme = customFormats
        Console.WriteLine("Custom theme applied to the entire presentation.")
    End Sub

    Private Sub CustomizeLayoutTheme(presentation As Presentation)
        ' Find target layout by name
        Dim targetLayoutName As String = "Title Slide Layout" ' Change to your layout name
        Dim targetLayout = presentation.SlideMasters _
            .SelectMany(Function(sm) sm.Layouts) _
            .FirstOrDefault(Function(l) l.Name = targetLayoutName)
        If targetLayout IsNot Nothing Then

            ' Create theme override for layout
            Dim customTheme As New ThemeOverride()
            Dim customColors As New ThemeColorScheme()
            customColors.Accent1 = New OfficeColor(Color.Purple)
            customColors.Accent2 = New OfficeColor(Color.Orange)
            customTheme.ColorScheme = customColors

            Dim overrideFonts As New ThemeFontScheme()
            overrideFonts.MajorFont = New ThemeFont("Times New Roman", "Comic Sans MS", "Comic Sans MS")
            overrideFonts.MinorFont = New ThemeFont("Comic Sans MS", "Comic Sans MS")
            customTheme.FontScheme = overrideFonts

            Dim customFormats As New ThemeFormatScheme()
            customFormats.Fills.ResetTo(New List(Of DevExpress.Docs.Office.Fill) From {
                New DevExpress.Docs.Office.SolidFill(Color.Purple),
                New DevExpress.Docs.Office.SolidFill(Color.Orange),
                New DevExpress.Docs.Office.SolidFill(Color.Yellow)
            })
            customTheme.FormatScheme = customFormats

            ' Apply theme override to all slides with the target layout
            For Each slide In presentation.Slides.Where(Function(s) s.Layout Is targetLayout)
                slide.ThemeOverride = customTheme
            Next
            Console.WriteLine($"Custom theme applied to all slides with layout '{targetLayoutName}'.")
        Else
            Console.WriteLine($"Layout '{targetLayoutName}' not found.")
        End If
    End Sub

    Private Sub ExportToPdfAndOpen(presentation As Presentation, pdfPath As String)
        Using pdfStream As New FileStream(pdfPath, FileMode.Create)
            presentation.ExportToPdf(pdfStream)
        End Using
        Console.WriteLine($"PDF exported: {pdfPath}")
        ' Open the PDF file in the default application
        Process.Start(New ProcessStartInfo With {
            .FileName = pdfPath,
            .UseShellExecute = True
        })
    End Sub

    Private Sub SavePresentationFile(presentation As Presentation, pptxPath As String)
        Using outputStream As New FileStream(pptxPath, FileMode.Create)
            presentation.SaveDocument(outputStream)
        End Using
        Console.WriteLine($"Presentation saved: {pptxPath}")
    End Sub
End Module
