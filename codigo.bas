Sub GerarSVGEPDF()
    Dim codigo As String
    Dim nome As String
    Dim referencia As String
    Dim origem As String
    Dim conteudo As String
    Dim numeroProjeto As String
    Dim data As String
    Dim nomeProjeto As String
    Dim dataFormatada As String
    Dim dimensoes As String
    Dim peso As String
    Dim svgTemplate As String
    Dim svgOutput As String
    Dim pdfOutput As String
    Dim filePath As String
    Dim folderPath As String
    Dim FSO As Object
    Dim fileOut As Object
    Dim inkScapePath As String
    Dim command As String

    ' Obtém os valores das células do arquivo Excel atual
    codigo = ThisWorkbook.Sheets(1).Range("B1").Value
    nome = ThisWorkbook.Sheets(1).Range("B2").Value
    referencia = ThisWorkbook.Sheets(1).Range("B3").Value
    origem = ThisWorkbook.Sheets(1).Range("B4").Value
    conteudo = ThisWorkbook.Sheets(1).Range("B5").Value
    numeroProjeto = ThisWorkbook.Sheets(1).Range("B6").Value
    data = Format(ThisWorkbook.Sheets(1).Range("B7").Value, "yyyy/mm/dd")
    nomeProjeto = ThisWorkbook.Sheets(1).Range("B6").Value & "_" & ThisWorkbook.Sheets(1).Range("B8").Value
    dimensoes = ThisWorkbook.Sheets(1).Range("B9").Value
    peso = ThisWorkbook.Sheets(1).Range("B10").Value
    dataFormatada = Format(ThisWorkbook.Sheets(1).Range("B7").Value, "yyyyMMdd")

    ' Define valores padrão e cor para os campos se estiverem em branco
    If dimensoes = "" Then
        dimensoes = "XXXxYYYxZZZ"
    End If
    If peso = "" Then
        peso = "XXX kg"
    End If

    ' Caminho do arquivo Excel e do template SVG
    filePath = "G:\Outros computadores\O meu Portátil\DOCS\JAVASCRIPT\VBA-Inkscape\Cartao Visita\Card Template.svg" ' Atualize com o caminho correto do template SVG

    ' Caminho para o Inkscape
    inkScapePath = """C:\Program Files\Inkscape\bin\inkscape.exe""" ' Caminho atualizado para Inkscape

    ' Verifica se o arquivo template existe
    If Dir(filePath) = "" Then
        MsgBox "Arquivo de template SVG não encontrado: " & filePath, vbCritical
        Exit Sub
    End If

    ' Lê o conteúdo do template SVG
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrorHandler
    svgTemplate = FSO.OpenTextFile(filePath, 1).ReadAll
    On Error GoTo 0

    ' Substitui os placeholders pelos valores do formulário
    svgTemplate = Replace(svgTemplate, "[Cod]", codigo)
    svgTemplate = Replace(svgTemplate, "[nome]", nome)
    svgTemplate = Replace(svgTemplate, "[referencia]", referencia)
    svgTemplate = Replace(svgTemplate, "[origem]", origem)
    svgTemplate = Replace(svgTemplate, "[conteudo]", conteudo)
    svgTemplate = Replace(svgTemplate, "[numero_projeto]", numeroProjeto)
    svgTemplate = Replace(svgTemplate, "[data]", data)
    svgTemplate = Replace(svgTemplate, "[nome_projeto]", nomeProjeto)
    svgTemplate = Replace(svgTemplate, "[data_formatada]", data)

    ' Adiciona o texto dimensoes e peso ao SVG
    If InStr(svgTemplate, "[dimensoes]") > 0 Then
        If dimensoes = "XXXxYYYxZZZ" Then
            svgTemplate = Replace(svgTemplate, "[dimensoes]", "<text id='dim_molde' x='...' y='...' fill='red'>" & dimensoes & "</text>")
        Else
            svgTemplate = Replace(svgTemplate, "[dimensoes]", "<text id='dim_molde' x='...' y='...' fill='black'>" & dimensoes & "</text>")
        End If
    End If

    If InStr(svgTemplate, "[peso]") > 0 Then
        If peso = "XXX kg" Then
            svgTemplate = Replace(svgTemplate, "[peso]", "<text id='peso_molde' x='...' y='...' fill='red'>" & peso & "</text>")
        Else
            svgTemplate = Replace(svgTemplate, "[peso]", "<text id='peso_molde' x='...' y='...' fill='black'>" & peso & "</text>")
        End If
    End If

    ' Define o caminho da pasta e dos arquivos de saída
    folderPath = "G:\Outros computadores\O meu Portátil\DOCS\JAVASCRIPT\VBA-Inkscape\Cartao Visita\" & nomeProjeto
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath
        If Err.Number <> 0 Then
            MsgBox "Erro ao criar o diretório: " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo ErrorHandler
    End If

    svgOutput = folderPath & "\" & codigo & "_Etiqueta_" & dataFormatada & ".svg"
    pdfOutput = folderPath & "\" & codigo & "_Etiqueta_" & dataFormatada & ".pdf"

    ' Usa ADODB.Stream para salvar o arquivo SVG com codificação UTF-8
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = "utf-8"
        .Open
        .WriteText svgTemplate
        .SaveToFile svgOutput, 2 ' 2 = adSaveCreateOverWrite
        .Close
    End With

    ' Converte o SVG para PDF e aplica a conversão em curvas usando Inkscape
    On Error GoTo ErrorHandler
    command = inkScapePath & " """ & svgOutput & """ --export-filename=""" & pdfOutput & """ --export-text-to-path"
    Shell command, vbNormalFocus

    ' Mensagem de sucesso
    MsgBox "A etiqueta " & codigo & " foi gerada com sucesso!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

