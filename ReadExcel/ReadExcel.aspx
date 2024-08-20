<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ReadExcel.aspx.cs" Inherits="ReadExcel.ReadExcel" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Solicitações em Lote</title>

    <script type="text/javascript">

        function excelInvalido() {
            Swal.fire('Atenção', 'Por favor, escolha um arquivo Excel válido.', 'warning');
        }

        function sucessoImportacao(insertCount) {
            Swal.fire('Sucesso', 'Foram processadas ' + insertCount + ' Solicitações.', 'success');
        }

        function excelInvalido() {
            Swal.fire('Atenção', 'Faça o upload de um arquivo Excel válido.', 'warning');
        }

        function excelError() {
            Swal.fire('Erro', 'Erro ao ler as linhas e colunas do Excel.', 'error');
        }

        function erroInsert(errorMessage) {
            Swal.fire('Erro', errorMessage, 'error');
        }
    </script>

    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script type="text/javascript" src="Scripts/bootstrap.min.js"></script>
    <script type="text/javascript" src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>

</head>
<body>
    <form id="form1" runat="server">
        <div class="row">
            <div class="col-md-5" style="margin-top: 1%;">
                <div class="input-group">
                    <asp:FileUpload ID="fileUpload" runat="server" type="file" CssClass="form-control" />
                    <div class="input-group-append">
                        <asp:Button ID="btnUpload" runat="server" Text="Carregar" CssClass="btn btn-primary" OnClick="btnUpload_Click" />
                    </div>
                </div>
            </div>
            <div class="col-md-5" style="margin-top: 1%;">
                <div class="d-flex">
                    <asp:Button ID="btnInsertData" runat="server" Text="Importar Solicitações" OnClick="btnInsertData_Click" CssClass="btn btn-primary" />
                </div>
            </div>
        </div>
        <br />
        <div id="resultPlaceholder" runat="server" style="text-align: center;"></div>
    </form>
</body>
</html>
