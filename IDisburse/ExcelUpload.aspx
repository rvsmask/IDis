<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ExcelUpload.aspx.cs" Inherits="IDisburse.ExcelUpload" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
        <div class="container mt-3">
    <label for="fileExcel" class="form-label">Upload Excel</label>
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
</asp:Content>
