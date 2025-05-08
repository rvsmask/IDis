<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ExcelUpload.aspx.cs" Inherits="Excel_Upload_ExcelUpload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excel Import</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">
</head>
<body>
    <form id="form1" runat="server">
        <div class="container mt-3">
            <label for="fileExcel" class="form-label" style="font-weight: bold;">Upload Excel</label>
            <div class="mb-3 d-flex justify-content-center">
                <div class="input-group has-validation">
                    <%--<span class="input-group-text" id="inputGroupPrepend">Choose Excel</span>--%>
                    <asp:FileUpload ID="fileExcel" runat="server" CssClass="form-control" aria-describedby="inputGroupPrepend" required />
                    <asp:Button ID="btnExclUpload" runat="server" OnClick="btnExclUpload_Click" Text="Upload File" CssClass="btn btn-outline-secondary" />
                    <div class="invalid-feedback">
                        Please choose a excel file
                    </div>
                </div>
            </div>
        </div>

        <div>
            <asp:GridView ShowHeaderWhenEmpty="true" ID="GridEmp" runat="server" AutoGenerateColumns="false" CssClass="table table-striped table-bordered">
                <HeaderStyle CssClass="kt-shape-bg-color-3" />
                <Columns>
                    <asp:TemplateField ControlStyle-CssClass="col-xs-1" HeaderText="Sr.No">
                        <ItemTemplate>
                            <asp:HiddenField ID="id" runat="server" Value="id" />
                            <span class="d-inline-block">
                                <%#Container.DataItemIndex + 1%>
                            </span>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="EmpName" HeaderText="Employee Name" SortExpression="EmpName" ItemStyle-CssClass="col-md-3" />
                    <asp:BoundField DataField="EmpDesignation" HeaderText="Emp Designation" SortExpression="EmpDesignation" ItemStyle-CssClass="col-md-4" />
                    <asp:BoundField DataField="EmpSalary" HeaderText="Emp Salary" SortExpression="EmpSalary" ItemStyle-CssClass="col-md-2" />
                </Columns>
            </asp:GridView>
        </div>

    </form>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js" integrity="sha384-KyZXEAg3QhqLMpG8r+J2Wk5vqXn3Fm/z2N1r8f6VZJ4T3Hdvh4kXG1j4fZ6IsU2f5" crossorigin="anonymous"></script>
</body>
</html>
