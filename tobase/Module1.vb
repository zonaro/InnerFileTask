Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports InnerLibs
Imports InnerLibs.LINQ
Imports Microsoft.Win32
Module Module1

    Sub Main()
        Try
            If FileParameters.Count > 0 Then
                Select Case Environment.GetCommandLineArgs()(1)
                    Case "--friendlyname"
                        If Confirm("Deseja realmente renomear estes arquivos?") Then
                            For Each f In FileParameters
                                If f.FullName.IsDirectoryPath Then
                                    For Each arquivo In Directory.GetFiles(f.FullName, "*", SearchOption.AllDirectories)
                                        FriendlyRename(New FileInfo(arquivo))
                                    Next
                                Else
                                    Dim arquivo = New FileInfo(f.FullName)
                                    FriendlyRename(arquivo)
                                End If
                            Next
                        End If
                    Case "--tobase64"
                        createbase()
                    Case "--copypath"
                        Notify("Copiando caminho de " & FileParameters(0).Name)
                        Clipboard.SetText(FileParameters(0).FullName)
                    Case "--enum"
                        Dim pat = Prompt("Digite o novo padrão de nome de arquivo: " & Environment.NewLine & Environment.NewLine & "Utilize o caractere # para a enumeração e o caractere $ para copiar o antigo nome do arquivo.")
                        If pat.IsNotBlank Then
                            If Not pat.Contains("#") Then
                                pat = pat & " (#)"
                            End If
                            Dim filenumber As Integer = 0
                            For Each f In FileParameters
                                If f.FullName.IsDirectoryPath Then
                                    For Each arquivo In Directory.GetFiles(f.FullName, "*", SearchOption.AllDirectories)
                                        filenumber.Increment
                                        EnumRename(New FileInfo(arquivo), pat, filenumber)
                                    Next
                                Else Dim arquivo = New FileInfo(f.FullName)
                                    filenumber.Increment
                                    EnumRename(arquivo, pat, filenumber)
                                End If
                            Next
                        End If
                    Case "--grayscale"
                        GrayScale()
                    Case "--combinevertical"
                        Combine(True)
                    Case "--combinehorizontal"
                        Combine(False)
                    Case "--cleanempty"
                        If Confirm("Essa operação vai eliminar todos os diretórios vazios." & Environment.NewLine & Environment.NewLine & "Deseja continuar?") Then
                            For Each dir As DirectoryInfo In FileParameters.Where(Function(p) p.FullName.IsDirectoryPath)
                                Notify("Realizando limpeza em " & dir.Name)
                                dir.CleanDirectory()
                            Next
                        End If
                    Case "--copycontent"
                        Dim txt As String = ""
                        Dim l As New List(Of Image)
                        For Each file As FileInfo In FileParameters.Where(Function(x) x.FullName.IsFilePath)
                            If New FileType(file).IsText Then
                                Notify("Copiando texto de " & file.Name.Quote)
                                txt &= IO.File.ReadAllText(file.FullName)
                                Continue For
                            End If
                            If New FileType(file).IsImage Then
                                Notify("Copiando imagem de " & file.Name.Quote)
                                l.Add(Image.FromFile(file.FullName))
                                Continue For
                            End If
                            If New FileType(file).IsAudio Then
                                Notify("Copiando audio de " & file.Name.Quote)
                                Clipboard.SetAudio(file.ToBytes)
                                Exit Sub
                            End If
                            Notify("Arquivo não suportado")
                        Next



                        If l.Count > 0 Then
                            Dim img As Image
                            If txt.IsNotBlank Then
                                l.Add(txt.DrawImage())
                            End If
                            If l.Count > 1 Then
                                img = CombineImages(l, Confirm("Deseja combinar as imagens verticalmente?"))
                            Else
                                img = l.First
                            End If
                            Clipboard.SetImage(img)
                            Exit Sub
                        End If
                        If txt.IsNotBlank Then
                            Clipboard.SetText(txt)
                        End If
                    Case "--watermark"
                        watermark()
                    Case "--crop"
                        crop()
                    Case "--circle"
                        circle()
                    Case Else
                        FileParameters.ForEach(Sub(b) Process.Start(b.FullName))
                End Select
            Else
                CreateOrDestroyShortcuts()
            End If
        Catch ex As Exception
            WinForms.Alert(ex.Message)
        End Try
        Application.Exit()
    End Sub

    Private Sub createbase()
        For Each file In FileParameters
            If file.FullName.IsFilePath Then
                Dim f = New FileInfo(file.FullName)
                Clipboard.SetText(f.ToDataURL)
                Notify(f.Name.Quote & " copiado como DataURL")
                Exit Sub
            End If
        Next
    End Sub

    ReadOnly Property FileParameters As List(Of FileSystemInfo)
        Get
            Dim f As New List(Of FileSystemInfo)
            For index = 2 To Environment.GetCommandLineArgs().Length - 1
                Dim p = Environment.GetCommandLineArgs()(index)
                Select Case True
                    Case p.IsDirectoryPath
                        f.Add(New DirectoryInfo(p))
                    Case p.IsFilePath
                        f.Add(New FileInfo(p))
                    Case Else
                End Select
            Next
            Return f
        End Get
    End Property

    Sub circle()
        Try
            Dim imagens As New List(Of Image)
            For Each file As FileInfo In FileParameters.Where(Function(x) x.FullName.IsFilePath AndAlso New FileType(Path.GetExtension(x.FullName)).IsImage)
                Dim novaimagem = Image.FromFile(file.FullName).CropToCircle
                Dim caminho = file.FullName.Replace(file.Name, Path.GetFileNameWithoutExtension(file.FullName) & "_circle." & Path.GetExtension(file.FullName).Trim("."))
                Notify("Aplicando corte de circulo em " & file.Name)
                novaimagem.Save(caminho, Imaging.ImageFormat.Png)
            Next
        Catch ex As Exception
            Alert("Erro ao cortar imagens")
        End Try
    End Sub

    Sub crop()
        Try
            Dim s = Prompt("Digite o tamanho de corte da imagem:", "200x200")
            Dim size = s.ToSize
            Dim imagens As New List(Of Image)
            For Each file As FileInfo In FileParameters.Where(Function(x) x.FullName.IsFilePath AndAlso New FileType(Path.GetExtension(x.FullName)).IsImage)
                Dim novaimagem = Image.FromFile(file.FullName).Crop(size)
                Dim caminho = file.FullName.Replace(file.Name, Path.GetFileNameWithoutExtension(file.FullName) & "_" & s.RemoveAny(Path.GetInvalidFileNameChars) & "." & Path.GetExtension(file.FullName).Trim("."))
                Notify("Aplicando corte em " & file.Name)
                novaimagem.Save(caminho, Imaging.ImageFormat.Png)
            Next
        Catch ex As Exception
            Alert("Erro ao cortar imagens")
        End Try
    End Sub

    Sub watermark()
        Dim imgmk As Image
        Dim mk = Prompt("Digite a a marca d'água ou o caminho da imagem que será utilizada:")
        If mk.IsNotBlank Then
            If mk.IsFilePath AndAlso New FileType(Path.GetExtension(mk)).IsImage Then
                imgmk = Image.FromFile(mk)
            Else
                imgmk = mk.DrawImage(Nothing, Color.Gray, Color.Transparent)
            End If

            For Each file As FileInfo In FileParameters.Where(Function(x) x.FullName.IsFilePath AndAlso New FileType(Path.GetExtension(x.FullName)).IsImage)
                Dim novaimagem = Image.FromFile(file.FullName)
                novaimagem = novaimagem.InsertWatermark(imgmk.Resize(novaimagem.Width, novaimagem.Height))
                Dim caminho = file.FullName.Replace(file.Name, Path.GetFileNameWithoutExtension(file.FullName) & "_wtmrk." & Path.GetExtension(file.FullName).Trim("."))
                Notify("Aplicando marca d'água em " & file.Name)
                novaimagem.Save(caminho, Imaging.ImageFormat.Png)
            Next
        End If

    End Sub
    Sub GrayScale()
        Dim imagens As New List(Of Image)
        For Each file As FileInfo In FileParameters.Where(Function(x) x.FullName.IsFilePath AndAlso New FileType(Path.GetExtension(x.FullName)).IsImage)
            Dim novaimagem = Image.FromFile(file.FullName).ConvertToGrayscale
            Dim caminho = file.FullName.Replace(file.Name, Path.GetFileNameWithoutExtension(file.FullName) & "_grayscale." & Path.GetExtension(file.FullName).Trim("."))
            Notify("Aplicando Grayscale em " & file.Name)
            novaimagem.Save(caminho, Imaging.ImageFormat.Png)
        Next
    End Sub

    Sub Combine(flow As Boolean)
        Dim imagens As New List(Of Image)
        Dim dir As DirectoryInfo
        For Each file In FileParameters
            If file.FullName.IsFilePath AndAlso New FileType(file.Extension).IsImage Then
                dir = New FileInfo(file.FullName).Directory
                imagens.Add(Image.FromFile(file.FullName))
                Notify("Adicionando imagem " & file.Name)
            End If
        Next
        If imagens.Count > 1 Then
            Dim novaimagem As Bitmap = imagens.CombineImages(flow)
            Dim caminho = dir.FullName & "\NovaCombinação.png"
            novaimagem.Save(caminho, Imaging.ImageFormat.Png)
        End If
    End Sub

    Sub FriendlyRename(Arquivo As FileInfo)
        Try
            Notify("Renomeando " & Arquivo.Name)
            My.Computer.FileSystem.RenameFile(Arquivo.FullName, Path.GetFileNameWithoutExtension(Arquivo.Name).ToFriendlyURL(True) & Arquivo.Extension)
        Catch ex As Exception
            Notify("Falha ao renomear " & Arquivo.Name & Environment.NewLine & ex.Message)
        End Try

    End Sub

    Sub EnumRename(Arquivo As FileInfo, Expression As String, Number As Integer)
        Try
            Notify("Renomeando " & Arquivo.Name)
            My.Computer.FileSystem.RenameFile(Arquivo.FullName, Expression.Replace("#", Number.ToString).Replace("$", Path.GetFileNameWithoutExtension(Arquivo.Name)) & Arquivo.Extension)
        Catch ex As Exception
            Notify("Falha ao renomear " & Arquivo.Name & Environment.NewLine & ex.Message)
        End Try
    End Sub

    Sub CreateIcons()
        Dim paths As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.SendTo))
        paths.CreateShortcut("InnerFileTask - Renomear como URL amigável", "--friendlyname")
        paths.CreateShortcut("InnerFileTask - Copiar como DataURL", "--tobase64")
        paths.CreateShortcut("InnerFileTask - Copiar caminho do arquivo", "--copypath")
        paths.CreateShortcut("InnerFileTask - Renomear e enumerar em massa", "--enum")
        paths.CreateShortcut("InnerFileTask - Combinar imagens verticalmente", "--combinevertical")
        paths.CreateShortcut("InnerFileTask - Combinar imagens horizontalmente", "--combinehorizontal")
        paths.CreateShortcut("InnerFileTask - Limpar diretórios vazios", "--cleanempty")
        paths.CreateShortcut("InnerFileTask - Copiar conteudo do arquivo", "--copycontent")
        paths.CreateShortcut("InnerFileTask - Converter imagem para preto e branco", "--grayscale")
        paths.CreateShortcut("InnerFileTask - Aplicar marca d'água", "--watermark")
        paths.CreateShortcut("InnerFileTask - Cortar imagens", "--crop")
        paths.CreateShortcut("InnerFileTask - Cortar imagens para circulo", "--circle")
    End Sub

    Sub CreateOrDestroyShortcuts()
        Dim paths As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.SendTo))
        Dim l = paths.EnumerateFiles.Where(Function(x) x.Name.StartsWith("InnerFileTask"))
        If l.Count > 0 Then
            If MsgBox("Deseja remover os atalhos do InnerFileTask?", vbInformation + vbYesNo) = vbYes Then
                l.ForEach(Sub(b) b.DeleteIfExist())
            Else
                CreateIcons()
            End If
        Else
            CreateIcons()
        End If
    End Sub



End Module
