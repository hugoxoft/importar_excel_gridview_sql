<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="impExcelaGridview.aspx.vb" Inherits="proyectos.impExcelaGridview" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div class="row">
            <div class="row">
                <asp:Label runat="server" Text="Se realizará la importación de el contenido de un archivo de Excel"></asp:Label>
                <br />
                <asp:Label runat="server" Text="a un GridView y posteriormete a la tabla de la base de datos"></asp:Label>
            </div>
            <br />
            <div class="row">
                <asp:FileUpload ID="FileUpload_archivo" runat="server" CssClass="form-control" />
            </div>
            <br />
            <div class="row">                
                <asp:Button ID="Button_visualizar" runat="server" Text="Visualizar registros" Width="150px" />
                <asp:Button ID="Button_limpiar" runat="server" Text="Limpiar visualización" Width="150px" />
                <asp:Button ID="Button_importar" runat="server" Text="Realizar importación" Width="150px" />
            </div>
            <br />
            <div class="row">
                <div class="col-md-12">
                    <div style="height: 500px; overflow: auto;">
                        <asp:Label ID="Label_aviso" runat="server" Font-Bold="True"></asp:Label>
                        <asp:GridView ID="GridView1" runat="server" class="table table-bordered table-hover" Width="100%">
                        </asp:GridView>
                    </div>
                </div>
            </div>
        </div>
    </form>
</body>
</html>
