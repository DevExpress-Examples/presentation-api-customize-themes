' See https://aka.ms/new-console-template for more information
Imports System.Drawing
Imports System.Diagnostics
Imports System.Linq
Imports DevExpress.Docs
Imports DevExpress.Docs.Presentation

Friend Class Program

    Shared Sub Main()
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
            If choice Is "0" Then Exit While
             ''' Cannot convert LocalDeclarationStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.RemodelVariableDeclaration(VariableDeclarationSyntax declaration) in C:\builds\CS2VB\CodeConverter-master\CodeConverter\VB\CommonConversions.cs:line 478
'''    at ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in C:\builds\CS2VB\CodeConverter-master\CodeConverter\VB\MethodBodyExecutableStatementVisitor.cs:line 59
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in C:\builds\CS2VB\CodeConverter-master\CodeConverter\VB\CommentConvertingMethodBodyVisitor.cs:line 24
''' 
''' Input:
''' 
'''             string fileSuffix = choice switch {
'''                 "1" => "colorscheme",
'''                 "2" => "fontscheme",
'''                 "3" => "formatscheme",
'''                 "4" => "shapestyle",
'''                 "5" => "themeoverride",
'''                 "6" => "mapping",
'''                 "7" => "customtheme",
'''                 "8" => "layouttheme",
'''                 _ => "result"
'''             };
''' 
'''  Dim pptxPath As String = Path.Combine(Environment.CurrentDirectory, $"presentation_{fileSuffix}.pptx")
            Dim pdfPath As String = Path.Combine(Environment.CurrentDirectory, $"presentation_{fileSuffix}.pdf")
            ' Open the presentation file
            Using fs As FileStream = File.OpenRead(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.pptx"))
                Dim presentation As Presentation = New Presentation(fs)
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
                Program.ExportToPdfAndOpen(presentation, pdfPath)
                Console.Write($"PDF will be saved as '{pdfPath}'. Do you also want to save the presentation file (pptx)? (y/n): ")
                Dim saveChoice = Console.ReadLine()
                If saveChoice?.Trim().ToLower() Is "y" Then
                    Program.SavePresentationFile(presentation, pptxPath)
                End If
            End Using

            Console.WriteLine()
        End While

        Console.WriteLine("Application exited.")
    End Sub

    Private Shared Sub CustomizeColorScheme(ByVal presentation As Presentation)
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

    Private Shared Sub CustomizeFontScheme(ByVal presentation As Presentation)
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

    Private Shared Sub CustomizeFormatScheme(ByVal presentation As Presentation)
        ' Get first slide and its master
        Dim slide = presentation.Slides(0)
        Dim mySlideMaster = presentation.SlideMasters.Where(Function(sm) sm.Layouts.Contains(slide.Layout)).FirstOrDefault()
        Dim myTheme = mySlideMaster?.Theme
        If myTheme Is Nothing Then Return
        ' Edit format scheme fills
        myTheme.FormatScheme.Fills.ResetTo(New List(Of Fill) From {New SolidFill(Color.Red), New SolidFill(Color.Magenta), New SolidFill(Color.Lime)})
        ' Edit format scheme background fills
        myTheme.FormatScheme.BackgroundFills.ResetTo(New List(Of Fill) From {New SolidFill(Color.Purple), New SolidFill(Color.Yellow), New SolidFill(Color.DarkGoldenrod)})
        ' Edit format scheme line styles
        myTheme.FormatScheme.LineStyles.ResetTo(New List(Of LineStyle) From {New LineStyle() With {.Width = 50, .DashType = LineDashType.Dot, .Fill = New SolidFill(Color.Red)}, New LineStyle() With {.Width = 20, .DashType = LineDashType.DashDot, .Fill = New SolidFill(Color.Magenta)}, New LineStyle() With {.Width = 5, .DashType = LineDashType.Solid, .Fill = New SolidFill(Color.Lime)}})
        ' Edit format scheme effects
        myTheme.FormatScheme.Effects.ResetTo(New List(Of ShapeEffectProperties) From {New ShapeEffectProperties() With {.Blur = New BlurEffect(100, True)}, New ShapeEffectProperties() With {.InnerShadow = New InnerShadowEffect()}, New ShapeEffectProperties() With {.Glow = New GlowEffect(15)}})
        Console.WriteLine("Format scheme customized for the presentation.")
    End Sub

    Private Shared Sub EditShapeStyle(ByVal presentation As Presentation)
        ' Edit style of the first shape on the first slide
        Dim slide = presentation.Slides(0)
        If slide.Shapes.Count Is 0 Then
            ' Add a rectangle shape if none exists
            Dim shape = New Shape(ShapeType.Rectangle) With {.X = 100, .Y = 100, .Width = 400, .Height = 200, .Style = New ShapeStyle() With {.FillThemeIndex = 1, .OutlineThemeIndex = 1}, .TextArea = New TextArea("Sample themed shape") With {.Level1ParagraphProperties = New ParagraphProperties() With {.TextProperties = New TextProperties() With {.FontSize = 32}}}}
            slide.Shapes.Add(shape)
        End If

        Dim targetShape = TryCast(slide.Shapes(0), Shape)
        If targetShape IsNot Nothing Then
            targetShape.Style.FillThemeIndex = 2 ' Use theme fill index
            targetShape.Style.OutlineThemeIndex = 3 ' Use theme outline index
            targetShape.Style.FontStyle = FontSchemeStyle.Major
            targetShape.Style.FontColor = New OfficeColor(SchemeColors.Accent3)
            ' You can also set direct fill/outline
            ' targetShape.Fill = new SolidFill(Color.LightGreen);
            ' targetShape.Outline = new Outline() { Width = 5, Fill = new SolidFill(Color.DarkBlue) };
            Console.WriteLine("Shape style edited on the first slide.")
        End If
    End Sub

    Private Shared Sub ApplyThemeOverride(ByVal presentation As Presentation)
        ' Get first slide
        Dim slide As Slide = presentation.Slides(0)
        ' Create theme override for the slide
        Dim themeOverride As ThemeOverride = New ThemeOverride()
        Dim overrideColors As ThemeColorScheme = New ThemeColorScheme()
        overrideColors.Accent1 = New OfficeColor(Color.Red)
        overrideColors.Accent2 = New OfficeColor(Color.Red)
        themeOverride.ColorScheme = overrideColors
        Dim formatScheme As ThemeFormatScheme = New ThemeFormatScheme()
        themeOverride.FormatScheme = formatScheme
        ' Set required fills
        themeOverride.FormatScheme.Fills.ResetTo(New List(Of Fill) From {New SolidFill(Color.Red), New SolidFill(Color.Magenta), New SolidFill(Color.Lime)})
        ' Apply override to slide
        slide.ThemeOverride = themeOverride
        Console.WriteLine("ThemeOverride applied to the first slide.")
    End Sub

    Private Shared Sub RemapColors(ByVal presentation As Presentation)
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
        Dim slideColorMap As ColorMap = New ColorMap()
        slideColorMap.Accent1 = ColorSchemeIndexType.Accent6
        slideColorMap.Hyperlink = ColorSchemeIndexType.Accent5
        slide.ColorMapOverride = slideColorMap
        Console.WriteLine("Master and second slide colors remapped.")
    End Sub

    Private Shared Sub ApplyCustomTheme(ByVal presentation As Presentation)
        ' Get first slide and its master
        Dim slide = presentation.Slides(0)
        Dim mySlideMaster = presentation.SlideMasters.Where(Function(sm) sm.Layouts.Contains(slide.Layout)).FirstOrDefault()
        Dim myTheme = mySlideMaster?.Theme
        If myTheme Is Nothing Then Return
        ' Create custom color scheme
        Dim customColors As ThemeColorScheme = New ThemeColorScheme()
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
        Dim customFonts As ThemeFontScheme = New ThemeFontScheme()
        customFonts.Name = "Corporate Fonts"
        customFonts.MajorFont = New ThemeFont("Segoe UI", "Arial", "Arial")
        customFonts.MinorFont = New ThemeFont("Calibri", "Times New Roman", "Times New Roman")
        ' Create custom format scheme
        Dim customFormats As ThemeFormatScheme = New ThemeFormatScheme()
        ' Add fills (minimum 3 required)
        customFormats.Fills.Add(New SolidFill(New OfficeColor(Color.FromArgb(68, 114, 196))))
        customFormats.Fills.Add(New GradientFill(GradientType.Linear, _(New GradientStop(New OfficeColor(Color.White), 0.0), New GradientStop(New OfficeColor(Color.FromArgb(68, 114, 196)), 1.0))))
        customFormats.Fills.Add(New SolidFill(New OfficeColor(Color.FromArgb(112, 173, 71))))
        ' Add background fills (minimum 3 required)
        customFormats.BackgroundFills.Add(New SolidFill(New OfficeColor(Color.White)))
        customFormats.BackgroundFills.Add(New SolidFill(New OfficeColor(Color.FromArgb(242, 246, 252))))
        customFormats.BackgroundFills.Add(New GradientFill(GradientType.Linear, _(New GradientStop(New OfficeColor(Color.White), 0.0), New GradientStop(New OfficeColor(Color.FromArgb(242, 246, 252)), 1.0))))
        ' Add line styles (minimum 3 required)
        customFormats.LineStyles.Add(New LineStyle() With {.Width = 1, .Fill = New SolidFill(New OfficeColor(Color.FromArgb(68, 114, 196)))})
        customFormats.LineStyles.Add(New LineStyle() With {.Width = 2, .Fill = New SolidFill(New OfficeColor(Color.FromArgb(112, 173, 71)))})
        customFormats.LineStyles.Add(New LineStyle() With {.Width = 3, .Fill = New SolidFill(New OfficeColor(Color.FromArgb(255, 192, 0)))})
        ' Add effects (minimum 3 required)
        customFormats.Effects.Add(New ShapeEffectProperties() With {.Blur = New BlurEffect(4, True)})
        customFormats.Effects.Add(New ShapeEffectProperties() With {.InnerShadow = New InnerShadowEffect()})
        customFormats.Effects.Add(New ShapeEffectProperties() With {.Glow = New GlowEffect(5)})
        ' Assign custom theme to master
        myTheme.ColorScheme = customColors
        myTheme.FontScheme = customFonts
        myTheme.FormatScheme = customFormats
        Console.WriteLine("Custom theme applied to the entire presentation.")
    End Sub

    Private Shared Sub CustomizeLayoutTheme(ByVal presentation As Presentation)
        ' Find target layout by name
        Dim targetLayoutName As String = "Title Slide Layout" ' Change to your layout name
        Dim targetLayout = presentation.SlideMasters.SelectMany(Function(sm) sm.Layouts).FirstOrDefault(Function(l) l.Name Is targetLayoutName)
        If targetLayout IsNot Nothing Then
            ' Create theme override for layout
            Dim customTheme As ThemeOverride = New ThemeOverride()
            Dim customColors As ThemeColorScheme = New ThemeColorScheme()
            customColors.Accent1 = New OfficeColor(Color.Purple)
            customColors.Accent2 = New OfficeColor(Color.Orange)
            customTheme.ColorScheme = customColors
            Dim overrideFonts As ThemeFontScheme = New ThemeFontScheme()
            overrideFonts.MajorFont = New ThemeFont("Times New Roman", "Comic Sans MS", "Comic Sans MS")
            overrideFonts.MinorFont = New ThemeFont("Comic Sans MS", "Comic Sans MS")
            customTheme.FontScheme = overrideFonts
            Dim customFormats As ThemeFormatScheme = New ThemeFormatScheme()
            customFormats.Fills.ResetTo(New List(Of Fill) From {New SolidFill(Color.Purple), New SolidFill(Color.Orange), New SolidFill(Color.Yellow)})
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

    Private Shared Sub ExportToPdfAndOpen(ByVal presentation As Presentation, ByVal pdfPath As String)
        Using pdfStream As FileStream = New FileStream(pdfPath, FileMode.Create)
            presentation.ExportToPdf(pdfStream)
        End Using

        Console.WriteLine($"PDF exported: {pdfPath}")
        ' Open the PDF file in the default application
        Process.Start(New ProcessStartInfo With {.FileName = pdfPath, .UseShellExecute = True})
    End Sub

    Private Shared Sub SavePresentationFile(ByVal presentation As Presentation, ByVal pptxPath As String)
        Using outputStream As FileStream = New FileStream(pptxPath, FileMode.Create)
            presentation.SaveDocument(outputStream)
        End Using

        Console.WriteLine($"Presentation saved: {pptxPath}")
    End Sub
End Class
