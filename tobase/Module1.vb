Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports InnerLibs
Imports Microsoft.Win32
Module Module1

    Sub Main()
        Try
            CreateIcons()
            Dim args = Environment.GetCommandLineArgs()
            If args.Length > 2 Then
                Select Case args(1)
                    Case "--friendlyname"
                        If Confirm("Deseja realmente renomear estes arquivos?") Then
                            For index = 2 To args.Length - 1
                                If File.GetAttributes(args(index)) = FileAttributes.Directory Then
                                    For Each arquivo In Directory.GetFiles(args(index), "*", SearchOption.AllDirectories)
                                        FriendlyRename(New FileInfo(arquivo))
                                    Next
                                Else
                                    Dim arquivo = New FileInfo(args(index))
                                    FriendlyRename(arquivo)
                                End If
                            Next
                        End If
                    Case "--tobase64"
                        If New FileInfo(args(2)).GetFileType.Contains("image") Then
                            Notify("Copiando Base64 de " & New FileInfo(args(2)).Name)
                            Clipboard.SetText(Image.FromFile(args(2)).ToDataURL())
                        Else
                            WinForms.Alert("Isso não pode ser convertido para DataURL, apenas imagens são permitidas.")
                        End If

                    Case "--copypath"
                        Notify("Copiando caminho de " & New FileInfo(args(2)).Name)
                        Clipboard.SetText(New FileInfo(args(2)).FullName)
                    Case "--enum"
                        Dim pat = Prompt("Digite o novo padrão de nome de arquivo: " & Environment.NewLine & Environment.NewLine & "Utilize o caractere # para a enumeração e o caractere $ para copiar o antigo nome do arquivo.")
                        If pat.IsNotBlank Then
                            If Not pat.Contains("#") Then
                                pat = pat & " (#)"
                            End If
                            Dim filenumber As Integer = 0
                            For index = 2 To args.Length - 1
                                If File.GetAttributes(args(index)) = FileAttributes.Directory Then
                                    For Each arquivo In Directory.GetFiles(args(index), "*", SearchOption.AllDirectories)
                                        filenumber.Increment
                                        EnumRename(New FileInfo(arquivo), pat, filenumber)
                                    Next
                                Else Dim arquivo = New FileInfo(args(index))
                                    filenumber.Increment
                                    EnumRename(arquivo, pat, filenumber)
                                End If
                            Next
                        End If
                    Case "--grayscale"
                        GrayScale(args)
                    Case "--combinevertical"
                        Combine(args, True)
                    Case "--combinehorizontal"
                        Combine(args, False)
                    Case "--cleanempty"
                        If Confirm("Essa operação vai eliminar todos os diretórios vazios." & Environment.NewLine & Environment.NewLine & "Deseja continuar?") Then
                            For index = 2 To args.Length - 1
                                If File.GetAttributes(args(index)) = FileAttributes.Directory Then
                                    Dim dir As New DirectoryInfo(args(index))
                                    Notify("Realizando limpeza em " & dir.Name)
                                    dir.CleanDirectory()
                                End If
                            Next
                        End If
                    Case "--copytext"
                        Dim txt As String = ""
                        For index = 2 To args.Length - 1
                            Notify("Copiando texto de " & args(index).Quote)
                            txt &= File.ReadAllText(args(index))
                        Next
                        If txt.IsNotBlank Then
                            Clipboard.SetText(txt)
                        End If
                    Case Else
                        Process.Start(New FileInfo(args(2)).FullName)
                End Select
            Else
                WinForms.Alert("Nenhuma tarefa selecionada.")
            End If
        Catch ex As Exception
            WinForms.Alert(ex.Message)
        End Try
        Application.Exit()
    End Sub



    Sub GrayScale(args As String())
        Dim imagens As New List(Of Image)
        For index = 2 To args.Length - 1
            Dim arq = New FileInfo(args(index))
            If arq.GetFileType().Contains("image") Then
                Dim novaimagem = Image.FromFile(args(index)).ConvertToGrayscale
                Dim caminho = arq.FullName.Replace(arq.Name, Path.GetFileNameWithoutExtension(arq.FullName) & "_grayscale." & Path.GetExtension(arq.FullName.Trim(".")))
                novaimagem.Save(caminho, Imaging.ImageFormat.Png)
                Notify("Aplicando Grayscale em " & arq.Name)
            End If
        Next
    End Sub

    Sub Combine(args As String(), flow As Boolean)
        Dim imagens As New List(Of Image)
        For index = 2 To args.Length - 1
            If New FileInfo(args(index)).GetFileType().Contains("image") Then
                imagens.Add(Image.FromFile(args(index)))
                Notify("Adicionando imagem " & New FileInfo(args(index)).Name)
            End If
        Next
        Dim novaimagem As Bitmap = imagens.CombineImages(flow)
        Clipboard.SetImage(novaimagem)
        Dim caminho = New FileInfo(args(2)).DirectoryName & "\NovaCombinação.jpg"
        novaimagem.Save(caminho, Imaging.ImageFormat.Jpeg)
    End Sub

    Sub FriendlyRename(Arquivo As FileInfo)
        Try
            Notify("Renomeando " & Arquivo.Name)
            My.Computer.FileSystem.RenameFile(Arquivo.FullName, Arquivo.Name.Split(".")(0).ToFriendlyURL(True) & Arquivo.Extension)
        Catch ex As Exception
            Notify("Falha ao renomear " & Arquivo.Name & Environment.NewLine & ex.Message)
        End Try

    End Sub

    Sub EnumRename(Arquivo As FileInfo, Expression As String, Number As Integer)
        Try
            Notify("Renomeando " & Arquivo.Name)
            My.Computer.FileSystem.RenameFile(Arquivo.FullName, Expression.Replace("#", Number.ToString).Replace("$", Arquivo.Name.Split(".")(0)) & Arquivo.Extension)
        Catch ex As Exception
            Notify("Falha ao renomear " & Arquivo.Name & Environment.NewLine & ex.Message)
        End Try
    End Sub

    Sub CreateIcons()
        Dim paths As New DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.SendTo))
        paths.CreateShortcut("InnerFileTask - Renomear como URL amigável", "--friendlyname")
        paths.CreateShortcut("InnerFileTask - Copiar como DataURL", "--tobase64")
        paths.CreateShortcut("InnerFileTask - Copiar caminho do arquivo", "--copypath")
        paths.CreateShortcut("InnerFileTask - Renomear ou enumerar em massa", "--enum")
        paths.CreateShortcut("InnerFileTask - Combinar imagens verticalmente", "--combinevertical")
        paths.CreateShortcut("InnerFileTask - Combinar imagens horizontalmente", "--combinehorizontal")
        paths.CreateShortcut("InnerFileTask - Limpar diretórios vazios", "--cleanempty")
        paths.CreateShortcut("InnerFileTask - Copiar Texto do Arquivo", "--copytext")
        paths.CreateShortcut("InnerFileTask - Converter imagem para preto e branco", "--grayscale")

    End Sub



End Module
