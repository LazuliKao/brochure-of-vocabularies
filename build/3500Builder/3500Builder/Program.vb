Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Module Program
    Sub Main(args As String())
        Const RawPath = "C:\Users\gaoxi\OneDrive\文档\study\hs\raw.txt"
        'Dim r As New Regex(
        '    "(?<word>.*?)\[(?<yb>.*?)\]\s*?(?<me>.*)",
        '                   RegexOptions.IgnorePatternWhitespace Or RegexOptions.ExplicitCapture Or RegexOptions.Compiled)
        'For Each l In File.ReadAllText(RawPath).Replace(vbCr, "").Split(vbLf)
        '    Dim t = l.Trim.Replace("a.", "adj.").Replace("ad.", "adv.")
        '    If t.Length = 1 Then
        '        'd.Paragraphs.Last.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
        '        'GetLastRange().InsertParagraphAfter()
        '        'd.Paragraphs.Last.LeftIndent = 0
        '        'd.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        '        'InsertWithFont(t, Sub(r, rf, rd)
        '        '                      r.Font.Size = 24
        '        '                  End Sub)
        '        Continue For 'A-Z
        '    Else
        '        Try
        '            'Dim m = r.Match(t)
        '            'If Not m.Success Then
        '            '    Console.WriteLine(t)
        '            '    Console.ReadKey()
        '            'End If
        '            'd.Paragraphs.Add()
        '            'GetLastRange().InsertParagraphAfter()
        '            'd.Paragraphs.Last.LeftIndent = 0
        '            'd.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        '            'd.Paragraphs.Last.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
        '            'd.Paragraphs.Last.LineSpacing = 12
        '            Dim iet = t.IndexOf("]") + 1
        '            Dim forward = t.Remove(iet)

        '            Dim ist = t.IndexOf("[")
        '            'InsertWithFont(forward.Remove(ist), Sub(r, rf, rd)
        '            '                                        r.Font.Name = "Comic Sans MS"
        '            '                                    End Sub)
        '            'InsertWithFont("/" & forward.Substring(ist).Replace("[", "").Replace("]", "") & "/",
        '            '               Sub(r, rf, rd)
        '            '                   r.Font.Color = Word.WdColor.wdColorGray30
        '            '                   r.Font.Name = "Arial"
        '            '               End Sub)
        '            'GetLastRange().Font.Name = "Comic Sans MS"
        '            'GetLastRange().InsertAfter(forward.Remove(ist))
        '            'GetLastRange().Font.Name = "Arial"
        '            'GetLastRange().InsertAfter(forward.Substring(ist))
        '            Dim back = t.Substring(iet)
        '            'For index = forward.Length To 40
        '            '    back = " " + back
        '            'Next
        '            'For index = 1 To 10
        '            '    back = vbTab + back
        '            'Next
        '            'GetLastRange().InsertParagraphAfter()
        '            'd.Paragraphs.Last.TabIndent(5)
        '            Try
        '                Do While back.Length > 0
        '                    Dim idot = back.IndexOf(".") + 1
        '                    If idot = 0 Then
        '                        Throw New Exception("aaaaaaaaaaa")
        '                    End If
        '                    Dim receive = back.Remove(idot)
        '                    Dim fi = Regex.Match(receive, "((adj|adv|pron|v(i|t)?|prep|n)\.)$")
        '                    If fi.Success Then
        '                        'InsertWithFont(receive.Remove(fi.Index), Sub(r, rf, rd)
        '                        '                                             r.Font.Color = Word.WdColor.wdColorOliveGreen
        '                        '                                             r.Font.Name = "宋体"
        '                        '                                             r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                        '                                         End Sub)
        '                        'InsertWithFont(receive.Substring(fi.Index), Sub(r, rf, rd)
        '                        '                                                r.Font.Color = Word.WdColor.wdColorDarkRed
        '                        '                                                r.Font.Name = "Consolas"
        '                        '                                                r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                        '                                            End Sub)
        '                    Else
        '                        'InsertWithFont(receive, Sub(r, rf, rd)
        '                        '                            r.Font.Color = Word.WdColor.wdColorBlack
        '                        '                            r.Font.Name = "宋体"
        '                        '                            r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                        '                        End Sub)
        '                    End If
        '                    back = back.Substring(idot)
        '                Loop
        '            Catch ex As Exception
        '                If ex.Message <> "aaaaaaaaaaa" Then
        '                    Console.WriteLine("aaaaaaaaaaa" & ex.ToString)
        '                    Console.WriteLine(t)
        '                    Console.ReadLine()
        '                End If

        '                'InsertWithFont(back, Sub(r, rf, rd)
        '                '                         r.Font.Color = Word.WdColor.wdColorBlack
        '                '                         r.Font.Name = "宋体"
        '                '                         r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        '                '                     End Sub)
        '            End Try
        '        Catch ex As Exception
        '            Console.WriteLine(ex.ToString)
        '            Console.WriteLine(t)
        '            Console.ReadLine()
        '        End Try
        '        Console.WriteLine(t)
        '    End If
        'Next
        Dim a As New Word.Application
        a.ShowMe()
        Dim d As New Word.Document
        Console.WriteLine("start")
        d.Activate()
        d.PageSetup.LeftMargin = 50
        d.PageSetup.RightMargin = 20
        d.PageSetup.TopMargin = 20
        d.PageSetup.BottomMargin = 20
        d.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA5
        'd.PageSetup.FooterDistance = 10
        'd.Footnotes.Item(0).Range.Text = "233"
        'd.Sections(1).Headers(0).Range.Text = "修改后的内容"
        'd.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
        'Dim oSection As Section
        'Dim oHF As HeaderFooter

        'd.Footnotes.Item(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range = "233"
        'For Each oSection In d.Sections.OfType(Of Word.Section)
        '    With oSection.Footers(Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage).PageNumbers
        '        .NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleArabicFullWidth
        '        .RestartNumberingAtSection = True
        '        With oSection.PageSetup
        '            '首页不同
        '            .DifferentFirstPageHeaderFooter = False
        '            '奇偶页不同
        '            .OddAndEvenPagesHeaderFooter = False
        '        End With
        '        .StartingNumber = 1
        '        '.Add(Word.WdPageNumberAlignment.wdAlignPageNumberCenter, Nothing).Alignment = Word.WdPageNumberAlignment.wdAlignPageNumberCenter
        '    End With
        'Next
        '    d.capSub QQ1722187970()
        'Dim oSection As Section
        '    Dim oHF As HeaderFooter
        '    Debug.Print Word.ActiveDocument.Sections.Count
        'Dim oDoc As Document
        'Set oDoc = Word.ActiveDocument
        'Dim oPN As PageNumber
        'With oDoc
        '    For Each oSection In .Sections
        '        With oSection
        '            With .PageSetup
        '                '首页不同
        '                .DifferentFirstPageHeaderFooter = False
        '                '奇偶页不同
        '                .OddAndEvenPagesHeaderFooter = False
        '            End With
        '            oHF = .Footers(wdHeaderFooterPrimary)
        '            With oHF.PageNumbers
        '                .NumberStyle = wdPageNumberStyleArabicFullWidth
        '                ' 不续前节 '
        '                .RestartNumberingAtSection = True
        '                '从5开始编号
        '                .StartingNumber = 5
        '                oPN = .Add
        '                With oPN
        '                    .Alignment = wdAlignPageNumberCenter
        '                End With
        '            End With
        '        End With
        '    Next
        'End With
        'End Sub
        Dim i = 0
        Dim GetLastRange = Function()
                               Try
                                   Return d.Range(d.Words.Last.End - 1, d.Words.Last.End - 1)
                               Catch ex As Exception
                                   Return d.Range(0, 0)
                               End Try
                           End Function
        Dim InsertWithFont = Sub(text As String, act As Action(Of Word.Range, Word.Range, Word.Range))
                                 text = text.Trim
                                 Dim start = d.Words.Last.End - 1
                                 Dim rf = GetLastRange()
                                 rf.InsertAfter(text)
                                 Dim r = d.Range(start, d.Words.Last.End - 1)
                                 act?.Invoke(r, rf, GetLastRange())
                             End Sub
        Dim all = File.ReadAllText(RawPath).Replace(vbCr, "").Split(vbLf)
        Dim al = all.Length
        Dim lastTime = Now
        Dim esqueue As New Queue(Of Double)
        For Each l In all
            Task.Run(Sub()
                         i += 1
                         Dim es = (Now - lastTime).TotalMilliseconds
                         esqueue.Enqueue(es)
                         If esqueue.Count > 5 Then
                             esqueue.Dequeue()
                         End If
                         Dim ts = (esqueue.ToArray.Average())
                         Console.Title = $"{i}/{al} - {i / al:P2} {Math.Round(ts)}ms/个  {Math.Round(1 / ts * 1000 * 60)}个/min"
                         lastTime = Now
                     End Sub)
            Dim t = l.Trim
            If t.Length = 1 Then
                'd.Paragraphs.Last.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                GetLastRange().InsertParagraphAfter()
                d.Paragraphs.Last.LeftIndent = 0
                d.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                InsertWithFont(t, Sub(r, rf, rd)
                                      r.Font.Size = 24
                                      d.Paragraphs.Last.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
                                  End Sub)
                Continue For 'A-Z
            Else
                Try
                    'Dim m = r.Match(t)
                    'If Not m.Success Then
                    '    Console.WriteLine(t)
                    '    Console.ReadKey()
                    'End If
                    'd.Paragraphs.Add()
                    GetLastRange().InsertParagraphAfter()
                    d.Paragraphs.Last.LeftIndent = 0
                    d.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                    d.Paragraphs.Last.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
                    d.Paragraphs.Last.LineSpacing = 10
                    Dim iet = t.IndexOf("]") + 1
                    Dim forward = t.Remove(iet)

                    Dim ist = t.IndexOf("[")
                    Const strXML = "<?xml version=""1.0"" standalone=""yes""?>
<?mso-application progid=""Word.Document""?>
<w:wordDocument xmlns:aml=""http://schemas.microsoft.com/aml/2001/core""
    xmlns:wpc=""http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas""
    xmlns:cx=""http://schemas.microsoft.com/office/drawing/2014/chartex""
    xmlns:cx1=""http://schemas.microsoft.com/office/drawing/2015/9/8/chartex""
    xmlns:cx2=""http://schemas.microsoft.com/office/drawing/2015/10/21/chartex""
    xmlns:cx3=""http://schemas.microsoft.com/office/drawing/2016/5/9/chartex""
    xmlns:cx4=""http://schemas.microsoft.com/office/drawing/2016/5/10/chartex""
    xmlns:cx5=""http://schemas.microsoft.com/office/drawing/2016/5/11/chartex""
    xmlns:cx6=""http://schemas.microsoft.com/office/drawing/2016/5/12/chartex""
    xmlns:cx7=""http://schemas.microsoft.com/office/drawing/2016/5/13/chartex""
    xmlns:cx8=""http://schemas.microsoft.com/office/drawing/2016/5/14/chartex""
    xmlns:dt=""uuid:C2F41010-65B3-11d1-A29F-00AA00C14882""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""
    xmlns:aink=""http://schemas.microsoft.com/office/drawing/2016/ink""
    xmlns:am3d=""http://schemas.microsoft.com/office/drawing/2017/model3d""
    xmlns:o=""urn:schemas-microsoft-com:office:office""
    xmlns:v=""urn:schemas-microsoft-com:vml""
    xmlns:w10=""urn:schemas-microsoft-com:office:word""
    xmlns:w=""http://schemas.microsoft.com/office/word/2003/wordml""
    xmlns:wx=""http://schemas.microsoft.com/office/word/2003/auxHint""
    xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml""
    xmlns:wsp=""http://schemas.microsoft.com/office/word/2003/wordml/sp2""
    xmlns:sl=""http://schemas.microsoft.com/schemaLibrary/2003/core"" w:macrosPresent=""no"" w:embeddedObjPresent=""no"" w:ocxPresent=""no"" xml:space=""preserve"">
    <w:ignoreSubtree w:val=""http://schemas.microsoft.com/office/word/2003/wordml/sp2""/>
    <o:DocumentProperties>
        <o:Version>16</o:Version>
    </o:DocumentProperties>
    <w:fonts>
        <w:defaultFonts w:ascii=""等线"" w:fareast=""等线"" w:h-ansi=""等线"" w:cs=""Times New Roman""/>
        <w:font w:name=""Times New Roman"">
            <w:panose-1 w:val=""02020603050405020304""/>
            <w:charset w:val=""00""/>
            <w:family w:val=""Roman""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""E0002EFF"" w:usb-1=""C000785B"" w:usb-2=""00000009"" w:usb-3=""00000000"" w:csb-0=""000001FF"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""Arial"">
            <w:panose-1 w:val=""020B0604020202020204""/>
            <w:charset w:val=""00""/>
            <w:family w:val=""Swiss""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""E0002EFF"" w:usb-1=""C000785B"" w:usb-2=""00000009"" w:usb-3=""00000000"" w:csb-0=""000001FF"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""宋体"">
            <w:altName w:val=""SimSun""/>
            <w:panose-1 w:val=""02010600030101010101""/>
            <w:charset w:val=""86""/>
            <w:family w:val=""auto""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""00000003"" w:usb-1=""288F0000"" w:usb-2=""00000016"" w:usb-3=""00000000"" w:csb-0=""00040001"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""MS Gothic"">
            <w:altName w:val=""ＭＳ ゴシック""/>
            <w:panose-1 w:val=""020B0609070205080204""/>
            <w:charset w:val=""80""/>
            <w:family w:val=""Modern""/>
            <w:pitch w:val=""fixed""/>
            <w:sig w:usb-0=""E00002FF"" w:usb-1=""6AC7FDFB"" w:usb-2=""08000012"" w:usb-3=""00000000"" w:csb-0=""0002009F"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""Cambria Math"">
            <w:panose-1 w:val=""02040503050406030204""/>
            <w:charset w:val=""00""/>
            <w:family w:val=""Roman""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""E00006FF"" w:usb-1=""420024FF"" w:usb-2=""02000000"" w:usb-3=""00000000"" w:csb-0=""0000019F"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""等线"">
            <w:altName w:val=""DengXian""/>
            <w:panose-1 w:val=""02010600030101010101""/>
            <w:charset w:val=""86""/>
            <w:family w:val=""auto""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""A00002BF"" w:usb-1=""38CF7CFA"" w:usb-2=""00000016"" w:usb-3=""00000000"" w:csb-0=""0004000F"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""Comic Sans MS"">
            <w:panose-1 w:val=""030F0702030302020204""/>
            <w:charset w:val=""00""/>
            <w:family w:val=""Script""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""00000287"" w:usb-1=""00000013"" w:usb-2=""00000000"" w:usb-3=""00000000"" w:csb-0=""0000009F"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""@MS Gothic"">
            <w:panose-1 w:val=""020B0609070205080204""/>
            <w:charset w:val=""80""/>
            <w:family w:val=""Modern""/>
            <w:pitch w:val=""fixed""/>
            <w:sig w:usb-0=""E00002FF"" w:usb-1=""6AC7FDFB"" w:usb-2=""08000012"" w:usb-3=""00000000"" w:csb-0=""0002009F"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""@宋体"">
            <w:panose-1 w:val=""02010600030101010101""/>
            <w:charset w:val=""86""/>
            <w:family w:val=""auto""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""00000003"" w:usb-1=""288F0000"" w:usb-2=""00000016"" w:usb-3=""00000000"" w:csb-0=""00040001"" w:csb-1=""00000000""/>
        </w:font>
        <w:font w:name=""@等线"">
            <w:panose-1 w:val=""02010600030101010101""/>
            <w:charset w:val=""86""/>
            <w:family w:val=""auto""/>
            <w:pitch w:val=""variable""/>
            <w:sig w:usb-0=""A00002BF"" w:usb-1=""38CF7CFA"" w:usb-2=""00000016"" w:usb-3=""00000000"" w:csb-0=""0004000F"" w:csb-1=""00000000""/>
        </w:font>
    </w:fonts>
    <w:styles>
        <w:versionOfBuiltInStylenames w:val=""7""/>
        <w:latentStyles w:defLockedState=""off"" w:latentStyleCount=""376"">
            <w:lsdException w:name=""Normal""/>
            <w:lsdException w:name=""heading 1""/>
            <w:lsdException w:name=""heading 2""/>
            <w:lsdException w:name=""heading 3""/>
            <w:lsdException w:name=""heading 4""/>
            <w:lsdException w:name=""heading 5""/>
            <w:lsdException w:name=""heading 6""/>
            <w:lsdException w:name=""heading 7""/>
            <w:lsdException w:name=""heading 8""/>
            <w:lsdException w:name=""heading 9""/>
            <w:lsdException w:name=""index 1""/>
            <w:lsdException w:name=""index 2""/>
            <w:lsdException w:name=""index 3""/>
            <w:lsdException w:name=""index 4""/>
            <w:lsdException w:name=""index 5""/>
            <w:lsdException w:name=""index 6""/>
            <w:lsdException w:name=""index 7""/>
            <w:lsdException w:name=""index 8""/>
            <w:lsdException w:name=""index 9""/>
            <w:lsdException w:name=""toc 1""/>
            <w:lsdException w:name=""toc 2""/>
            <w:lsdException w:name=""toc 3""/>
            <w:lsdException w:name=""toc 4""/>
            <w:lsdException w:name=""toc 5""/>
            <w:lsdException w:name=""toc 6""/>
            <w:lsdException w:name=""toc 7""/>
            <w:lsdException w:name=""toc 8""/>
            <w:lsdException w:name=""toc 9""/>
            <w:lsdException w:name=""Normal Indent""/>
            <w:lsdException w:name=""footnote text""/>
            <w:lsdException w:name=""annotation text""/>
            <w:lsdException w:name=""header""/>
            <w:lsdException w:name=""footer""/>
            <w:lsdException w:name=""index heading""/>
            <w:lsdException w:name=""caption""/>
            <w:lsdException w:name=""table of figures""/>
            <w:lsdException w:name=""envelope address""/>
            <w:lsdException w:name=""envelope return""/>
            <w:lsdException w:name=""footnote reference""/>
            <w:lsdException w:name=""annotation reference""/>
            <w:lsdException w:name=""line number""/>
            <w:lsdException w:name=""page number""/>
            <w:lsdException w:name=""endnote reference""/>
            <w:lsdException w:name=""endnote text""/>
            <w:lsdException w:name=""table of authorities""/>
            <w:lsdException w:name=""macro""/>
            <w:lsdException w:name=""toa heading""/>
            <w:lsdException w:name=""List""/>
            <w:lsdException w:name=""List Bullet""/>
            <w:lsdException w:name=""List Number""/>
            <w:lsdException w:name=""List 2""/>
            <w:lsdException w:name=""List 3""/>
            <w:lsdException w:name=""List 4""/>
            <w:lsdException w:name=""List 5""/>
            <w:lsdException w:name=""List Bullet 2""/>
            <w:lsdException w:name=""List Bullet 3""/>
            <w:lsdException w:name=""List Bullet 4""/>
            <w:lsdException w:name=""List Bullet 5""/>
            <w:lsdException w:name=""List Number 2""/>
            <w:lsdException w:name=""List Number 3""/>
            <w:lsdException w:name=""List Number 4""/>
            <w:lsdException w:name=""List Number 5""/>
            <w:lsdException w:name=""Title""/>
            <w:lsdException w:name=""Closing""/>
            <w:lsdException w:name=""Signature""/>
            <w:lsdException w:name=""Default Paragraph Font""/>
            <w:lsdException w:name=""Body Text""/>
            <w:lsdException w:name=""Body Text Indent""/>
            <w:lsdException w:name=""List Continue""/>
            <w:lsdException w:name=""List Continue 2""/>
            <w:lsdException w:name=""List Continue 3""/>
            <w:lsdException w:name=""List Continue 4""/>
            <w:lsdException w:name=""List Continue 5""/>
            <w:lsdException w:name=""Message Header""/>
            <w:lsdException w:name=""Subtitle""/>
            <w:lsdException w:name=""Salutation""/>
            <w:lsdException w:name=""Date""/>
            <w:lsdException w:name=""Body Text First Indent""/>
            <w:lsdException w:name=""Body Text First Indent 2""/>
            <w:lsdException w:name=""Note Heading""/>
            <w:lsdException w:name=""Body Text 2""/>
            <w:lsdException w:name=""Body Text 3""/>
            <w:lsdException w:name=""Body Text Indent 2""/>
            <w:lsdException w:name=""Body Text Indent 3""/>
            <w:lsdException w:name=""Block Text""/>
            <w:lsdException w:name=""Hyperlink""/>
            <w:lsdException w:name=""FollowedHyperlink""/>
            <w:lsdException w:name=""Strong""/>
            <w:lsdException w:name=""Emphasis""/>
            <w:lsdException w:name=""Document Map""/>
            <w:lsdException w:name=""Plain Text""/>
            <w:lsdException w:name=""E-mail Signature""/>
            <w:lsdException w:name=""HTML Top of Form""/>
            <w:lsdException w:name=""HTML Bottom of Form""/>
            <w:lsdException w:name=""Normal (Web)""/>
            <w:lsdException w:name=""HTML Acronym""/>
            <w:lsdException w:name=""HTML Address""/>
            <w:lsdException w:name=""HTML Cite""/>
            <w:lsdException w:name=""HTML Code""/>
            <w:lsdException w:name=""HTML Definition""/>
            <w:lsdException w:name=""HTML Keyboard""/>
            <w:lsdException w:name=""HTML Preformatted""/>
            <w:lsdException w:name=""HTML Sample""/>
            <w:lsdException w:name=""HTML Typewriter""/>
            <w:lsdException w:name=""HTML Variable""/>
            <w:lsdException w:name=""Normal Table""/>
            <w:lsdException w:name=""annotation subject""/>
            <w:lsdException w:name=""No List""/>
            <w:lsdException w:name=""Outline List 1""/>
            <w:lsdException w:name=""Outline List 2""/>
            <w:lsdException w:name=""Outline List 3""/>
            <w:lsdException w:name=""Table Simple 1""/>
            <w:lsdException w:name=""Table Simple 2""/>
            <w:lsdException w:name=""Table Simple 3""/>
            <w:lsdException w:name=""Table Classic 1""/>
            <w:lsdException w:name=""Table Classic 2""/>
            <w:lsdException w:name=""Table Classic 3""/>
            <w:lsdException w:name=""Table Classic 4""/>
            <w:lsdException w:name=""Table Colorful 1""/>
            <w:lsdException w:name=""Table Colorful 2""/>
            <w:lsdException w:name=""Table Colorful 3""/>
            <w:lsdException w:name=""Table Columns 1""/>
            <w:lsdException w:name=""Table Columns 2""/>
            <w:lsdException w:name=""Table Columns 3""/>
            <w:lsdException w:name=""Table Columns 4""/>
            <w:lsdException w:name=""Table Columns 5""/>
            <w:lsdException w:name=""Table Grid 1""/>
            <w:lsdException w:name=""Table Grid 2""/>
            <w:lsdException w:name=""Table Grid 3""/>
            <w:lsdException w:name=""Table Grid 4""/>
            <w:lsdException w:name=""Table Grid 5""/>
            <w:lsdException w:name=""Table Grid 6""/>
            <w:lsdException w:name=""Table Grid 7""/>
            <w:lsdException w:name=""Table Grid 8""/>
            <w:lsdException w:name=""Table List 1""/>
            <w:lsdException w:name=""Table List 2""/>
            <w:lsdException w:name=""Table List 3""/>
            <w:lsdException w:name=""Table List 4""/>
            <w:lsdException w:name=""Table List 5""/>
            <w:lsdException w:name=""Table List 6""/>
            <w:lsdException w:name=""Table List 7""/>
            <w:lsdException w:name=""Table List 8""/>
            <w:lsdException w:name=""Table 3D effects 1""/>
            <w:lsdException w:name=""Table 3D effects 2""/>
            <w:lsdException w:name=""Table 3D effects 3""/>
            <w:lsdException w:name=""Table Contemporary""/>
            <w:lsdException w:name=""Table Elegant""/>
            <w:lsdException w:name=""Table Professional""/>
            <w:lsdException w:name=""Table Subtle 1""/>
            <w:lsdException w:name=""Table Subtle 2""/>
            <w:lsdException w:name=""Table Web 1""/>
            <w:lsdException w:name=""Table Web 2""/>
            <w:lsdException w:name=""Table Web 3""/>
            <w:lsdException w:name=""Balloon Text""/>
            <w:lsdException w:name=""Table Grid""/>
            <w:lsdException w:name=""Table Theme""/>
            <w:lsdException w:name=""Placeholder Text""/>
            <w:lsdException w:name=""No Spacing""/>
            <w:lsdException w:name=""Light Shading""/>
            <w:lsdException w:name=""Light List""/>
            <w:lsdException w:name=""Light Grid""/>
            <w:lsdException w:name=""Medium Shading 1""/>
            <w:lsdException w:name=""Medium Shading 2""/>
            <w:lsdException w:name=""Medium List 1""/>
            <w:lsdException w:name=""Medium List 2""/>
            <w:lsdException w:name=""Medium Grid 1""/>
            <w:lsdException w:name=""Medium Grid 2""/>
            <w:lsdException w:name=""Medium Grid 3""/>
            <w:lsdException w:name=""Dark List""/>
            <w:lsdException w:name=""Colorful Shading""/>
            <w:lsdException w:name=""Colorful List""/>
            <w:lsdException w:name=""Colorful Grid""/>
            <w:lsdException w:name=""Light Shading Accent 1""/>
            <w:lsdException w:name=""Light List Accent 1""/>
            <w:lsdException w:name=""Light Grid Accent 1""/>
            <w:lsdException w:name=""Medium Shading 1 Accent 1""/>
            <w:lsdException w:name=""Medium Shading 2 Accent 1""/>
            <w:lsdException w:name=""Medium List 1 Accent 1""/>
            <w:lsdException w:name=""Revision""/>
            <w:lsdException w:name=""List Paragraph""/>
            <w:lsdException w:name=""Quote""/>
            <w:lsdException w:name=""Intense Quote""/>
            <w:lsdException w:name=""Medium List 2 Accent 1""/>
            <w:lsdException w:name=""Medium Grid 1 Accent 1""/>
            <w:lsdException w:name=""Medium Grid 2 Accent 1""/>
            <w:lsdException w:name=""Medium Grid 3 Accent 1""/>
            <w:lsdException w:name=""Dark List Accent 1""/>
            <w:lsdException w:name=""Colorful Shading Accent 1""/>
            <w:lsdException w:name=""Colorful List Accent 1""/>
            <w:lsdException w:name=""Colorful Grid Accent 1""/>
            <w:lsdException w:name=""Light Shading Accent 2""/>
            <w:lsdException w:name=""Light List Accent 2""/>
            <w:lsdException w:name=""Light Grid Accent 2""/>
            <w:lsdException w:name=""Medium Shading 1 Accent 2""/>
            <w:lsdException w:name=""Medium Shading 2 Accent 2""/>
            <w:lsdException w:name=""Medium List 1 Accent 2""/>
            <w:lsdException w:name=""Medium List 2 Accent 2""/>
            <w:lsdException w:name=""Medium Grid 1 Accent 2""/>
            <w:lsdException w:name=""Medium Grid 2 Accent 2""/>
            <w:lsdException w:name=""Medium Grid 3 Accent 2""/>
            <w:lsdException w:name=""Dark List Accent 2""/>
            <w:lsdException w:name=""Colorful Shading Accent 2""/>
            <w:lsdException w:name=""Colorful List Accent 2""/>
            <w:lsdException w:name=""Colorful Grid Accent 2""/>
            <w:lsdException w:name=""Light Shading Accent 3""/>
            <w:lsdException w:name=""Light List Accent 3""/>
            <w:lsdException w:name=""Light Grid Accent 3""/>
            <w:lsdException w:name=""Medium Shading 1 Accent 3""/>
            <w:lsdException w:name=""Medium Shading 2 Accent 3""/>
            <w:lsdException w:name=""Medium List 1 Accent 3""/>
            <w:lsdException w:name=""Medium List 2 Accent 3""/>
            <w:lsdException w:name=""Medium Grid 1 Accent 3""/>
            <w:lsdException w:name=""Medium Grid 2 Accent 3""/>
            <w:lsdException w:name=""Medium Grid 3 Accent 3""/>
            <w:lsdException w:name=""Dark List Accent 3""/>
            <w:lsdException w:name=""Colorful Shading Accent 3""/>
            <w:lsdException w:name=""Colorful List Accent 3""/>
            <w:lsdException w:name=""Colorful Grid Accent 3""/>
            <w:lsdException w:name=""Light Shading Accent 4""/>
            <w:lsdException w:name=""Light List Accent 4""/>
            <w:lsdException w:name=""Light Grid Accent 4""/>
            <w:lsdException w:name=""Medium Shading 1 Accent 4""/>
            <w:lsdException w:name=""Medium Shading 2 Accent 4""/>
            <w:lsdException w:name=""Medium List 1 Accent 4""/>
            <w:lsdException w:name=""Medium List 2 Accent 4""/>
            <w:lsdException w:name=""Medium Grid 1 Accent 4""/>
            <w:lsdException w:name=""Medium Grid 2 Accent 4""/>
            <w:lsdException w:name=""Medium Grid 3 Accent 4""/>
            <w:lsdException w:name=""Dark List Accent 4""/>
            <w:lsdException w:name=""Colorful Shading Accent 4""/>
            <w:lsdException w:name=""Colorful List Accent 4""/>
            <w:lsdException w:name=""Colorful Grid Accent 4""/>
            <w:lsdException w:name=""Light Shading Accent 5""/>
            <w:lsdException w:name=""Light List Accent 5""/>
            <w:lsdException w:name=""Light Grid Accent 5""/>
            <w:lsdException w:name=""Medium Shading 1 Accent 5""/>
            <w:lsdException w:name=""Medium Shading 2 Accent 5""/>
            <w:lsdException w:name=""Medium List 1 Accent 5""/>
            <w:lsdException w:name=""Medium List 2 Accent 5""/>
            <w:lsdException w:name=""Medium Grid 1 Accent 5""/>
            <w:lsdException w:name=""Medium Grid 2 Accent 5""/>
            <w:lsdException w:name=""Medium Grid 3 Accent 5""/>
            <w:lsdException w:name=""Dark List Accent 5""/>
            <w:lsdException w:name=""Colorful Shading Accent 5""/>
            <w:lsdException w:name=""Colorful List Accent 5""/>
            <w:lsdException w:name=""Colorful Grid Accent 5""/>
            <w:lsdException w:name=""Light Shading Accent 6""/>
            <w:lsdException w:name=""Light List Accent 6""/>
            <w:lsdException w:name=""Light Grid Accent 6""/>
            <w:lsdException w:name=""Medium Shading 1 Accent 6""/>
            <w:lsdException w:name=""Medium Shading 2 Accent 6""/>
            <w:lsdException w:name=""Medium List 1 Accent 6""/>
            <w:lsdException w:name=""Medium List 2 Accent 6""/>
            <w:lsdException w:name=""Medium Grid 1 Accent 6""/>
            <w:lsdException w:name=""Medium Grid 2 Accent 6""/>
            <w:lsdException w:name=""Medium Grid 3 Accent 6""/>
            <w:lsdException w:name=""Dark List Accent 6""/>
            <w:lsdException w:name=""Colorful Shading Accent 6""/>
            <w:lsdException w:name=""Colorful List Accent 6""/>
            <w:lsdException w:name=""Colorful Grid Accent 6""/>
            <w:lsdException w:name=""Subtle Emphasis""/>
            <w:lsdException w:name=""Intense Emphasis""/>
            <w:lsdException w:name=""Subtle Reference""/>
            <w:lsdException w:name=""Intense Reference""/>
            <w:lsdException w:name=""Book Title""/>
            <w:lsdException w:name=""Bibliography""/>
            <w:lsdException w:name=""TOC Heading""/>
            <w:lsdException w:name=""Plain Table 1""/>
            <w:lsdException w:name=""Plain Table 2""/>
            <w:lsdException w:name=""Plain Table 3""/>
            <w:lsdException w:name=""Plain Table 4""/>
            <w:lsdException w:name=""Plain Table 5""/>
            <w:lsdException w:name=""Grid Table Light""/>
            <w:lsdException w:name=""Grid Table 1 Light""/>
            <w:lsdException w:name=""Grid Table 2""/>
            <w:lsdException w:name=""Grid Table 3""/>
            <w:lsdException w:name=""Grid Table 4""/>
            <w:lsdException w:name=""Grid Table 5 Dark""/>
            <w:lsdException w:name=""Grid Table 6 Colorful""/>
            <w:lsdException w:name=""Grid Table 7 Colorful""/>
            <w:lsdException w:name=""Grid Table 1 Light Accent 1""/>
            <w:lsdException w:name=""Grid Table 2 Accent 1""/>
            <w:lsdException w:name=""Grid Table 3 Accent 1""/>
            <w:lsdException w:name=""Grid Table 4 Accent 1""/>
            <w:lsdException w:name=""Grid Table 5 Dark Accent 1""/>
            <w:lsdException w:name=""Grid Table 6 Colorful Accent 1""/>
            <w:lsdException w:name=""Grid Table 7 Colorful Accent 1""/>
            <w:lsdException w:name=""Grid Table 1 Light Accent 2""/>
            <w:lsdException w:name=""Grid Table 2 Accent 2""/>
            <w:lsdException w:name=""Grid Table 3 Accent 2""/>
            <w:lsdException w:name=""Grid Table 4 Accent 2""/>
            <w:lsdException w:name=""Grid Table 5 Dark Accent 2""/>
            <w:lsdException w:name=""Grid Table 6 Colorful Accent 2""/>
            <w:lsdException w:name=""Grid Table 7 Colorful Accent 2""/>
            <w:lsdException w:name=""Grid Table 1 Light Accent 3""/>
            <w:lsdException w:name=""Grid Table 2 Accent 3""/>
            <w:lsdException w:name=""Grid Table 3 Accent 3""/>
            <w:lsdException w:name=""Grid Table 4 Accent 3""/>
            <w:lsdException w:name=""Grid Table 5 Dark Accent 3""/>
            <w:lsdException w:name=""Grid Table 6 Colorful Accent 3""/>
            <w:lsdException w:name=""Grid Table 7 Colorful Accent 3""/>
            <w:lsdException w:name=""Grid Table 1 Light Accent 4""/>
            <w:lsdException w:name=""Grid Table 2 Accent 4""/>
            <w:lsdException w:name=""Grid Table 3 Accent 4""/>
            <w:lsdException w:name=""Grid Table 4 Accent 4""/>
            <w:lsdException w:name=""Grid Table 5 Dark Accent 4""/>
            <w:lsdException w:name=""Grid Table 6 Colorful Accent 4""/>
            <w:lsdException w:name=""Grid Table 7 Colorful Accent 4""/>
            <w:lsdException w:name=""Grid Table 1 Light Accent 5""/>
            <w:lsdException w:name=""Grid Table 2 Accent 5""/>
            <w:lsdException w:name=""Grid Table 3 Accent 5""/>
            <w:lsdException w:name=""Grid Table 4 Accent 5""/>
            <w:lsdException w:name=""Grid Table 5 Dark Accent 5""/>
            <w:lsdException w:name=""Grid Table 6 Colorful Accent 5""/>
            <w:lsdException w:name=""Grid Table 7 Colorful Accent 5""/>
            <w:lsdException w:name=""Grid Table 1 Light Accent 6""/>
            <w:lsdException w:name=""Grid Table 2 Accent 6""/>
            <w:lsdException w:name=""Grid Table 3 Accent 6""/>
            <w:lsdException w:name=""Grid Table 4 Accent 6""/>
            <w:lsdException w:name=""Grid Table 5 Dark Accent 6""/>
            <w:lsdException w:name=""Grid Table 6 Colorful Accent 6""/>
            <w:lsdException w:name=""Grid Table 7 Colorful Accent 6""/>
            <w:lsdException w:name=""List Table 1 Light""/>
            <w:lsdException w:name=""List Table 2""/>
            <w:lsdException w:name=""List Table 3""/>
            <w:lsdException w:name=""List Table 4""/>
            <w:lsdException w:name=""List Table 5 Dark""/>
            <w:lsdException w:name=""List Table 6 Colorful""/>
            <w:lsdException w:name=""List Table 7 Colorful""/>
            <w:lsdException w:name=""List Table 1 Light Accent 1""/>
            <w:lsdException w:name=""List Table 2 Accent 1""/>
            <w:lsdException w:name=""List Table 3 Accent 1""/>
            <w:lsdException w:name=""List Table 4 Accent 1""/>
            <w:lsdException w:name=""List Table 5 Dark Accent 1""/>
            <w:lsdException w:name=""List Table 6 Colorful Accent 1""/>
            <w:lsdException w:name=""List Table 7 Colorful Accent 1""/>
            <w:lsdException w:name=""List Table 1 Light Accent 2""/>
            <w:lsdException w:name=""List Table 2 Accent 2""/>
            <w:lsdException w:name=""List Table 3 Accent 2""/>
            <w:lsdException w:name=""List Table 4 Accent 2""/>
            <w:lsdException w:name=""List Table 5 Dark Accent 2""/>
            <w:lsdException w:name=""List Table 6 Colorful Accent 2""/>
            <w:lsdException w:name=""List Table 7 Colorful Accent 2""/>
            <w:lsdException w:name=""List Table 1 Light Accent 3""/>
            <w:lsdException w:name=""List Table 2 Accent 3""/>
            <w:lsdException w:name=""List Table 3 Accent 3""/>
            <w:lsdException w:name=""List Table 4 Accent 3""/>
            <w:lsdException w:name=""List Table 5 Dark Accent 3""/>
            <w:lsdException w:name=""List Table 6 Colorful Accent 3""/>
            <w:lsdException w:name=""List Table 7 Colorful Accent 3""/>
            <w:lsdException w:name=""List Table 1 Light Accent 4""/>
            <w:lsdException w:name=""List Table 2 Accent 4""/>
            <w:lsdException w:name=""List Table 3 Accent 4""/>
            <w:lsdException w:name=""List Table 4 Accent 4""/>
            <w:lsdException w:name=""List Table 5 Dark Accent 4""/>
            <w:lsdException w:name=""List Table 6 Colorful Accent 4""/>
            <w:lsdException w:name=""List Table 7 Colorful Accent 4""/>
            <w:lsdException w:name=""List Table 1 Light Accent 5""/>
            <w:lsdException w:name=""List Table 2 Accent 5""/>
            <w:lsdException w:name=""List Table 3 Accent 5""/>
            <w:lsdException w:name=""List Table 4 Accent 5""/>
            <w:lsdException w:name=""List Table 5 Dark Accent 5""/>
            <w:lsdException w:name=""List Table 6 Colorful Accent 5""/>
            <w:lsdException w:name=""List Table 7 Colorful Accent 5""/>
            <w:lsdException w:name=""List Table 1 Light Accent 6""/>
            <w:lsdException w:name=""List Table 2 Accent 6""/>
            <w:lsdException w:name=""List Table 3 Accent 6""/>
            <w:lsdException w:name=""List Table 4 Accent 6""/>
            <w:lsdException w:name=""List Table 5 Dark Accent 6""/>
            <w:lsdException w:name=""List Table 6 Colorful Accent 6""/>
            <w:lsdException w:name=""List Table 7 Colorful Accent 6""/>
            <w:lsdException w:name=""Mention""/>
            <w:lsdException w:name=""Smart Hyperlink""/>
            <w:lsdException w:name=""Hashtag""/>
            <w:lsdException w:name=""Unresolved Mention""/>
            <w:lsdException w:name=""Smart Link""/>
        </w:latentStyles>
        <w:style w:type=""paragraph"" w:default=""on"" w:styleId=""Normal"">
            <w:name w:val=""Normal""/>
            <w:pPr>
                <w:widowControl w:val=""off""/>
                <w:jc w:val=""both""/>
            </w:pPr>
            <w:rPr>
                <wx:font wx:val=""等线""/>
                <w:kern w:val=""2""/>
                <w:sz w:val=""21""/>
                <w:sz-cs w:val=""22""/>
                <w:lang w:val=""EN-US"" w:fareast=""ZH-CN"" w:bidi=""AR-SA""/>
            </w:rPr>
        </w:style>
        <w:style w:type=""character"" w:default=""on"" w:styleId=""DefaultParagraphFont"">
            <w:name w:val=""Default Paragraph Font""/>
        </w:style>
        <w:style w:type=""table"" w:default=""on"" w:styleId=""TableNormal"">
            <w:name w:val=""Normal Table""/>
            <wx:uiName wx:val=""Table Normal""/>
            <w:rPr>
                <wx:font wx:val=""等线""/>
                <w:lang w:val=""EN-US"" w:fareast=""ZH-CN"" w:bidi=""AR-SA""/>
            </w:rPr>
            <w:tblPr>
                <w:tblInd w:w=""0"" w:type=""dxa""/>
                <w:tblCellMar>
                    <w:top w:w=""0"" w:type=""dxa""/>
                    <w:left w:w=""108"" w:type=""dxa""/>
                    <w:bottom w:w=""0"" w:type=""dxa""/>
                    <w:right w:w=""108"" w:type=""dxa""/>
                </w:tblCellMar>
            </w:tblPr>
        </w:style>
        <w:style w:type=""list"" w:default=""on"" w:styleId=""NoList"">
            <w:name w:val=""No List""/>
        </w:style>
    </w:styles>
    <w:shapeDefaults>
        <o:shapedefaults v:ext=""edit"" spidmax=""1027"" style=""mso-position-horizontal-relative:margin;mso-height-relative:margin"" fillcolor=""white"" stroke=""f"">
            <v:fill color=""white""/>
            <v:stroke on=""f""/>
        </o:shapedefaults>
        <o:shapelayout v:ext=""edit"">
            <o:idmap v:ext=""edit"" data=""1""/>
        </o:shapelayout>
    </w:shapeDefaults>
    <w:docPr>
        <w:view w:val=""print""/>
        <w:zoom w:percent=""173""/>
        <w:doNotEmbedSystemFonts/>
        <w:bordersDontSurroundHeader/>
        <w:bordersDontSurroundFooter/>
        <w:defaultTabStop w:val=""420""/>
        <w:drawingGridHorizontalSpacing w:val=""105""/>
        <w:drawingGridVerticalSpacing w:val=""156""/>
        <w:displayHorizontalDrawingGridEvery w:val=""0""/>
        <w:displayVerticalDrawingGridEvery w:val=""2""/>
        <w:punctuationKerning/>
        <w:characterSpacingControl w:val=""CompressPunctuation""/>
        <w:optimizeForBrowser/>
        <w:allowPNG/>
        <w:validateAgainstSchema/>
        <w:saveInvalidXML w:val=""off""/>
        <w:ignoreMixedContent w:val=""off""/>
        <w:alwaysShowPlaceholderText w:val=""off""/>
        <w:compat>
            <w:spaceForUL/>
            <w:balanceSingleByteDoubleByteWidth/>
            <w:doNotLeaveBackslashAlone/>
            <w:ulTrailSpace/>
            <w:doNotExpandShiftReturn/>
            <w:adjustLineHeightInTable/>
            <w:breakWrappedTables/>
            <w:snapToGridInCell/>
            <w:wrapTextWithPunct/>
            <w:useAsianBreakRules/>
            <w:dontGrowAutofit/>
            <w:useFELayout/>
        </w:compat>
        <wsp:rsids>
            <wsp:rsidRoot wsp:val=""004D242B""/>
            <wsp:rsid wsp:val=""002C4502""/>
            <wsp:rsid wsp:val=""004D242B""/>
            <wsp:rsid wsp:val=""009E58C9""/>
            <wsp:rsid wsp:val=""00D4475F""/>
        </wsp:rsids>
    </w:docPr>
    <w:body>
        <wx:sect>
            <w:p wsp:rsidR=""00D4475F"" wsp:rsidRDefault=""00D4475F"" wsp:rsidP=""009E58C9"">
                <w:pPr>
                    <w:ind w:left=""2100""/>
                    <w:jc w:val=""right""/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:noProof/>
                    </w:rPr>
                    <w:pict>
                        <v:shapetype id=""_x0000_t202"" coordsize=""21600,21600"" o:spt=""202"" path=""m,l,21600r21600,l21600,xe"">
                            <v:stroke joinstyle=""miter""/>
                            <v:path gradientshapeok=""t"" o:connecttype=""rect""/>
                        </v:shapetype>
                        <v:shape id=""Text Box 2"" o:spid=""_x0000_s1026"" type=""#_x0000_t202"" style=""position:absolute;left:0;text-align:left;margin-left:0;margin-top:-4pt;width:346.2pt;height:30pt;z-index:1;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:1pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:left;mso-position-horizontal-relative:margin;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top"" o:gfxdata=""UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF&#xA;90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA&#xA;0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD&#xA;OlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893&#xA;SUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y&#xA;JsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl&#xA;bHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR&#xA;JVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY&#xA;22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i&#xA;OWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA&#xA;IQDHaRjkCwIAAPQDAAAOAAAAZHJzL2Uyb0RvYy54bWysU9tuGyEQfa/Uf0C817ve2HG8Mo7SpKkq&#xA;pRcp6QdglvWiAkMBe9f9+g6s41jNW1QeEDAzZ+acGVbXg9FkL31QYBmdTkpKpBXQKLtl9OfT/Ycr&#xA;SkLktuEarGT0IAO9Xr9/t+pdLSvoQDfSEwSxoe4do12Mri6KIDppeJiAkxaNLXjDI179tmg87xHd&#xA;6KIqy8uiB984D0KGgK93o5GuM37bShG/t22QkWhGsbaYd5/3TdqL9YrXW89dp8SxDP6GKgxXFpOe&#xA;oO545GTn1Ssoo4SHAG2cCDAFtK0SMnNANtPyHzaPHXcyc0FxgjvJFP4frPi2/+GJahitpgtKLDfY&#xA;pCc5RPIRBlIlfXoXanR7dOgYB3zGPmeuwT2A+BWIhduO26288R76TvIG65umyOIsdMQJCWTTf4UG&#xA;0/BdhAw0tN4k8VAOgujYp8OpN6kUgY+zi+XlYoYmgbaL5bysxhS8fo52PsTPEgxJB0Y99j6j8/1D&#xA;iKkaXj+7pGQW7pXWuf/akp7R5bya54Azi1ERx1Mrw+hVmdY4MInkJ9vk4MiVHs+YQNsj60R0pByH&#xA;zYCOSYoNNAfk72EcQ/w2eOjA/6GkxxFkNPzecS8p0V8sariczhLhmC+z+aLCiz+3bM4t3AqEYjRS&#xA;Mh5vY57zkesNat2qLMNLJcdacbSyOsdvkGb3/J69Xj7r+i8AAAD//wMAUEsDBBQABgAIAAAAIQAz&#xA;0y682QAAAAUBAAAPAAAAZHJzL2Rvd25yZXYueG1sTI/NTsMwEITvSLyDtUjcqE0VIppmUyEQVxDl&#xA;R+LmxtskaryOYrcJb8/2BLdZzWrmm3Iz+16daIxdYITbhQFFXAfXcYPw8f58cw8qJsvO9oEJ4Yci&#xA;bKrLi9IWLkz8RqdtapSEcCwsQpvSUGgd65a8jYswEIu3D6O3Sc6x0W60k4T7Xi+NybW3HUtDawd6&#xA;bKk+bI8e4fNl//2Vmdfmyd8NU5iNZr/SiNdX88MaVKI5/T3DGV/QoRKmXTiyi6pHkCEJYZmBEjNf&#xA;ncVORJaDrkr9n776BQAA//8DAFBLAQItABQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAAAAAAAAA&#xA;AAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhADj9If/WAAAAlAEAAAsA&#xA;AAAAAAAAAAAAAAAALwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAMdpGOQLAgAA9AMAAA4A&#xA;AAAAAAAAAAAAAAAALgIAAGRycy9lMm9Eb2MueG1sUEsBAi0AFAAGAAgAAAAhADPTLrzZAAAABQEA&#xA;AA8AAAAAAAAAAAAAAAAAZQQAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAABAAEAPMAAABrBQAAAAA=&#xA;"" filled=""f"" stroked=""f"">
                            <v:textbox>
                                <w:txbxContent>
                                    <w:p wsp:rsidR=""009E58C9"" wsp:rsidRDefault=""009E58C9"" wsp:rsidP=""009E58C9"">
                                        <w:pPr>
                                            <w:spacing w:line=""200"" w:line-rule=""exact""/>
                                            <w:jc w:val=""left""/>
                                            <w:rPr>
                                                <w:rFonts w:ascii=""Arial"" w:h-ansi=""Arial"" w:cs=""Arial""/>
                                                <wx:font wx:val=""Arial""/>
                                                <w:color w:val=""808000""/>
                                            </w:rPr>
                                        </w:pPr>
                                        <w:r wsp:rsidRPr=""009D3D96"">
                                            <w:rPr>
                                                <w:rFonts w:ascii=""Comic Sans MS"" w:h-ansi=""Comic Sans MS""/>
                                                <wx:font wx:val=""Comic Sans MS""/>
                                            </w:rPr>
                                            <w:t>wordhere </w:t>
                                        </w:r>
                                        <w:r wsp:rsidRPr=""009D3D96"">
                                            <w:rPr>
                                                <w:rFonts w:ascii=""Arial"" w:h-ansi=""Arial"" w:cs=""Arial""/>
                                                <wx:font wx:val=""Arial""/>
                                                <w:color w:val=""82BDEE""/>
                                            </w:rPr>
                                            <w:t>/</w:t>
                                        </w:r>
                                        <w:r wsp:rsidRPr=""009D3D96"">
                                            <w:rPr>
                                                <w:rFonts w:ascii=""Arial"" w:fareast=""MS Gothic"" w:h-ansi=""Arial"" w:cs=""Arial"" w:hint=""fareast""/>
                                                <wx:font wx:val=""MS Gothic""/>
                                                <w:color w:val=""8FC36B""/>
                                            </w:rPr>
                                            <w:t>gAAAOEBAAATAAAAAAAAAAAA</w:t>
                                        </w:r>
                                        <w:r wsp:rsidRPr=""009D3D96"">
                                            <w:rPr>
                                                <w:rFonts w:ascii=""Arial"" w:h-ansi=""Arial"" w:cs=""Arial""/>
                                                <wx:font wx:val=""Arial""/>
                                                <w:color w:val=""82BDEE""/>
                                            </w:rPr>
                                            <w:t>/</w:t>
                                        </w:r>
                                    </w:p>
                                </w:txbxContent>
                            </v:textbox>
                            <w10:wrap anchorx=""margin""/>
                        </v:shape>
                    </w:pict>
                </w:r>
                <w:r wsp:rsidR=""009E58C9"" wsp:rsidRPr=""004D242B"">
                    <w:rPr>
                        <w:rFonts w:ascii=""宋体"" w:fareast=""宋体"" w:h-ansi=""宋体""/>
                        <wx:font wx:val=""宋体""/>
                        <w:color w:val=""000000""/>
                    </w:rPr>
                    <w:t></w:t>
                </w:r>
            </w:p>
            <w:sectPr wsp:rsidR=""00D4475F"" wsp:rsidSect=""004D242B"">
                <w:pgSz w:w=""8391"" w:h=""11906"" w:code=""11""/>
                <w:pgMar w:top=""720"" w:right=""720"" w:bottom=""720"" w:left=""720"" w:header=""851"" w:footer=""992"" w:gutter=""0""/>
                <w:cols w:space=""425""/>
                <w:docGrid w:type=""lines"" w:line-pitch=""312""/>
            </w:sectPr>
        </wx:sect>
    </w:body>
</w:wordDocument>
"
                    GetLastRange().InsertXML(strXML _
                        .Replace("wordhere", forward.Remove(ist).Trim) _
                        .Replace("gAAAOEBAAATAAAAAAAAAAAA", forward.Substring(ist).Replace("[", "").Replace("]", "").Trim)
                        )
                    d.Paragraphs.Last.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly
                    d.Paragraphs.Last.LineSpacing = 10
                    'InsertWithFont(forward.Remove(ist), Sub(r, rf, rd)
                    '                                        r.Font.Name = "Comic Sans MS"
                    '                                    End Sub)
                    'InsertWithFont("/" & forward.Substring(ist).Replace("[", "").Replace("]", "") & "/",
                    '                   Sub(r, rf, rd)
                    '                       r.Font.Color = Word.WdColor.wdColorDarkYellow
                    '                       r.Font.Name = "Arial"
                    '                   End Sub)

                    'GetLastRange().Font.Name = "Comic Sans MS"
                    'GetLastRange().InsertAfter(forward.Remove(ist))
                    'GetLastRange().Font.Name = "Arial"
                    'GetLastRange().InsertAfter(forward.Substring(ist))
                    Dim back = t.Substring(iet)
                    'For index = forward.Length To 40
                    '    back = " " + back
                    'Next
                    'For index = 1 To 10
                    '    back = vbTab + back
                    'Next
                    'GetLastRange().InsertParagraphAfter()
                    d.Paragraphs.Last.TabIndent(3)
                    Dim FllDefault = Sub(texttmp As String)
                                         Dim text = texttmp
                                         Try
                                             Do While text.Length > 0
                                                 Dim m = Regex.Match(text, "（.*?）|\(.*?\)")
                                                 If m.Success Then
                                                     If m.Index > 0 Then
                                                         InsertWithFont(text.Remove(m.Index), Sub(r, rf, rd)
                                                                                                  r.Font.Color = Word.WdColor.wdColorBlack
                                                                                                  r.Font.Name = "宋体"
                                                                                                  r.Font.Size = 10
                                                                                                  r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                                                                              End Sub)
                                                     End If
                                                     InsertWithFont(text.Substring(m.Index, m.Length), Sub(r, rf, rd)
                                                                                                           r.Font.Color = Word.WdColor.wdColorTeal
                                                                                                           r.Font.Name = "宋体"
                                                                                                           r.Font.Size = 10
                                                                                                           r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                                                                                       End Sub)
                                                     Try
                                                         text = text.Substring(m.Index + m.Length)
                                                     Catch ex As Exception
                                                         text = ""
                                                     End Try
                                                 Else
                                                     Exit Do
                                                 End If
                                             Loop
                                             If text.Length > 0 Then
                                                 InsertWithFont(text, Sub(r, rf, rd)
                                                                          r.Font.Color = Word.WdColor.wdColorBlack
                                                                          r.Font.Name = "宋体"
                                                                          r.Font.Size = 10
                                                                          r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                                                      End Sub)
                                             End If
                                         Catch ex As Exception
                                             Console.WriteLine(ex.ToString)
                                         End Try
                                         'Try
                                         '    Do While Text.Length > 0
                                         '        Dim ipst = Text.IndexOfAny({"("c, "（"c})
                                         '        If ipst = -1 Then
                                         '            InsertWithFont(Text, Sub(r, rf, rd)
                                         '                                     r.Font.Color = Word.WdColor.wdColorBlack
                                         '                                     r.Font.Name = "宋体"
                                         '                                     r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                         '                                 End Sub)
                                         '            Exit Do
                                         '        Else
                                         '            Dim ipstend = Text.IndexOfAny({")"c, "）"c}) + 1
                                         '            InsertWithFont(Text.Remove(ipst), Sub(r, rf, rd)
                                         '                                                  r.Font.Color = Word.WdColor.wdColorBlack
                                         '                                                  r.Font.Name = "宋体"
                                         '                                                  r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                         '                                              End Sub)
                                         '            Dim tmp = Text
                                         '            If ipstend > Text.Length Then
                                         '                tmp = Text.Remove(ipstend)
                                         '                InsertWithFont(tmp.Substring(ipstend), Sub(r, rf, rd)
                                         '                                                           r.Font.Color = Word.WdColor.wdColorTeal
                                         '                                                           r.Font.Name = "宋体"
                                         '                                                           r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                         '                                                       End Sub)
                                         '                Text = Text.Substring(ipstend)
                                         '            Else
                                         '                Throw New Exception
                                         '            End If
                                         '        End If
                                         '    Loop
                                         'Catch ex As Exception
                                         '    Console.WriteLine(ex.ToString)
                                         '    InsertWithFont(text, Sub(r, rf, rd)
                                         '                             r.Font.Color = Word.WdColor.wdColorBlack
                                         '                             r.Font.Name = "宋体"
                                         '                             r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                         '                         End Sub)
                                         'End Try
                                     End Sub
                    Try
                        Do While back.Length > 0
                            Dim idot = back.IndexOf(".") + 1
                            If idot = 0 Then
                                Throw New Exception
                            End If
                            Dim receive = back.Remove(idot)
                            Dim fi = Regex.Match(receive, "((a|ad(j|v)?|pron|v(i|t)?|prep|n|conj|num|int|aux|pl|mod)\.)$")
                            If fi.Success Then
                                FllDefault(receive.Remove(fi.Index))
                                InsertWithFont(receive.Substring(fi.Index) _
                                    .Replace("adj.", "xxxxxxxxxxx.").Replace("adv.", "xxxxxxxxxxxxx.") _
                                    .Replace("a.", "adj.").Replace("ad.", "adv.") _
                                    .Replace("xxxxxxxxxxx.", "adj.").Replace("xxxxxxxxxxxxx.", "adv."),
                                               Sub(r, rf, rd)
                                                   r.Font.Color = Word.WdColor.wdColorOrange
                                                   r.Font.Name = "Consolas"
                                                   r.Font.Size = 9.5
                                                   r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                               End Sub)
                            Else
                                FllDefault(receive)
                            End If
                            back = back.Substring(idot)
                        Loop
                    Catch ex As Exception
                        FllDefault(back)
                    End Try
                Catch ex As Exception
                    InsertWithFont(t, Sub(r, rf, rd)
                                          r.Font.Color = Word.WdColor.wdColorBlack
                                          r.Font.Name = "宋体"
                                          r.Paragraphs.Last.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                                      End Sub)
                    Console.WriteLine(ex.ToString)
                    Console.WriteLine(t)
                    'Console.ReadLine()
                End Try
            End If
        Next
    End Sub
End Module
