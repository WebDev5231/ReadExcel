<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ReadRow.aspx.cs" Inherits="ReadExcel.ReadRow" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Detalhes da Linha</title>
    <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.0/dist/JsBarcode.all.min.js"></script>
    <link href="Content/bootstrap.min.css" rel="stylesheet" />
    <script type="text/javascript" src="Scripts/bootstrap.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <br />
            <table border="1">
                <tr>
                    <td id="detailsCell" runat="server" colspan="2"></td>
                </tr>
            </table>
            <br />
        </div>
        <div id="barcodeDiv"></div>
    </form>
    <script type="text/javascript">
        document.addEventListener("DOMContentLoaded", function () {
            var rowData = '<%= Request.QueryString["data"] %>';
            var columnValues = rowData.split(',');

            var barcodeDiv = document.getElementById("barcodeDiv");

            columnValues.forEach(function (columnValue) {
                var parts = columnValue.split(':');
                if (parts.length === 2) {
                    var barcodeValue = parts[1].trim();

                    if (!isNaN(barcodeValue.charAt(0))) {
                        var barcodeElement = document.createElement("img");
                        barcodeDiv.appendChild(barcodeElement);

                        JsBarcode(barcodeElement, barcodeValue, {
                            format: "CODE128",
                            displayValue: true,
                            fontSize: 14
                        });
                    }
                }
            });
        });
    </script>

</body>
</html>
