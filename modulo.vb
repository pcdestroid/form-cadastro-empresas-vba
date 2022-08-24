Private Sub CommandButton1_Click()
If TextBox1.Value <> "" And TextBox2.Value <> "" And TextBox4.Value <> "" And TextBox11.Value <> "" And TextBox9.Value <> "" And TextBox10.Value <> "" And TextBox7.Value <> "" And TextBox6.Value <> "" And TextBox4.Value <> "" And TextBox5.Value <> "" Then
Set cadastroForn = Workbooks.Open(Filename:="G:\Drives compartilhados\Compras\CADASTROS\FORNECEDORES\planilha.xlsx", ReadOnly:=False)
Set cadastro = cadastroForn.Sheets(1)
linha = cadastro.Cells(cadastro.Rows.Count, "A").End(xlUp).Row + 1
cadastro.Range("A" & linha).Value = "#" & linha
cadastro.Range("B" & linha).Value = TextBox1.Value 'CNPJ
cadastro.Range("C" & linha).Value = TextBox2.Value 'Razão Social
cadastro.Range("D" & linha).Value = 1
cadastro.Range("E" & linha).Value = "" 'Produto/Serviço
cadastro.Range("F" & linha).Value = TextBox4.Value 'Cidade
cadastro.Range("G" & linha).Value = TextBox11.Value 'Telefone
cadastro.Range("H" & linha).Value = TextBox9.Value 'Contato
cadastro.Range("I" & linha).Value = TextBox10.Value 'Email
cadastro.Range("J" & linha).Value = TextBox7.Value 'Endereço
cadastro.Range("K" & linha).Value = TextBox6.Value 'Bairro
cadastro.Range("L" & linha).Value = TextBox4.Value 'Cidade
cadastro.Range("M" & linha).Value = TextBox5.Value 'UF
cadastro.Range("N" & linha).Value = TextBox3.Value 'CEP
cadastro.Range("O" & linha).Value = "A" 'Situação
cadastro.Range("P" & linha).Value = "" 'Faturamento Minimo
cadastro.Range("Q" & linha).Value = Application.UserName 'Cadastrado Por
cadastroForn.Close SaveChanges:=True
Unload UserForm1
'ThisWorkbook.Close SaveChanges:=False
End If
End Sub

Private Sub TextBox1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo fim
Dim jsonObject As Object, item As Object
Dim objHTTP As Object
Dim cnpj As String, nome As String, cidade As String, bairro As String, numero As String, logradouro As String, cep As String, dados() As String

cnpj = Replace(Replace(Replace(TextBox1.Value, ".", ""), "/", ""), "-", "")

'Criamos nosso objeto de requisção e enviamos
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")        'CNPJ
URL = "https://thecollector.linkana.com/companies?cnpj=eq." & cnpj & "&limit=1"
objHTTP.Open "GET", URL, False
objHTTP.Send
strResult = objHTTP.responseText

'Depois ajustar o código
'#########################################################

palavra = """razao_social"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
nome = texto

palavra = """municipio"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
cidade = texto

palavra = """bairro"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
bairro = texto

palavra = """logradouro"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
logradouro = texto

palavra = """numero"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
numero = texto

palavra = """cep"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
cep = texto

palavra = """uf"""
posicao = InStr(1, strResult, palavra) + Len(palavra) + 2
texto = Mid(strResult, posicao, Len(strResult))
posicao = InStr(1, texto, ",")
texto = Mid(texto, 1, posicao - 2)
uf = texto

'#########################################################

If nome <> "details:null" Then: TextBox2.Value = nome
TextBox4.Value = cidade
TextBox7.Value = logradouro
TextBox6.Value = bairro
TextBox3.Value = cep
TextBox5.Value = uf
TextBox8.Value = numero
fim:

End Sub
Private Sub UserForm_Initialize()
'ThisWorkbook.Windows.Application.Visible = False
End Sub
Private Sub UserForm_Terminate()
'ThisWorkbook.Windows.Application.Visible = True
'ThisWorkbook.Close SaveChanges:=False
End Sub
