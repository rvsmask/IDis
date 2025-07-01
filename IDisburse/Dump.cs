<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="WebForm2.aspx.cs" Inherits="WebApplication5.WebForm2" %>
 
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">

      
        <asp:UpdatePanel ID="updpMain" runat="server" UpdateMode="Conditional">
            <ContentTemplate>
                <div class="container mt-4">
                    <h2>Workday Black Knight Map</h2>

                    <div class="row mb-3">
                        <div class="col-md-4">
                            <label for="drpGroup">Group:</label>
                            <asp:DropDownList ID="drpGroup" runat="server" CssClass="form-control" />
                        </div>
                        <div class="col-md-4 mt-4">
                            <asp:Button ID="btnGetRecords" runat="server" CssClass="btn btn-primary" Text="Get Records" OnClick="btnGetRecords_Click" />
                        </div>
                    </div>

                    <asp:ListView ID="Records" runat="server" OnItemCommand="Records_ItemCommand"  OnItemDeleting="Records_ItemDeleting">
                        <LayoutTemplate>
                            <table class="table table-bordered">
                                <thead>
                                    <tr>
                                        <th>Group</th>
                                        <th>Field</th>
                                        <th>Transaction</th>
                                        <th>Investor</th>
                                        <th>Action</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr runat="server" id="itemPlaceholder"></tr>
                                </tbody>
                            </table>
                        </LayoutTemplate>
                        <ItemTemplate>
                            <tr>
                                <td><%# Eval("MPGROUP") %></td>
                                <td><%# Eval("MPFIELD") %></td>
                                <td><%# Eval("MPTRAN") %></td>
                                <td><%# Eval("MPITVN") %></td>
                                <td>
                               <asp:LinkButton ID="LinkButton2" runat="server" CssClass="btn btn-danger btn-sm realDelete" UseSubmitBehavior="false"
    CommandName="MyDelete"
    CommandArgument='<%# Eval("MPGROUP") + "," + Eval("MPFIELD") + "," + Eval("MPTRAN") + "," + Eval("MPITVN") %>'
    OnClientClick="return openConfirmModal(this);">
    Delete
</asp:LinkButton>

                                 </td>
                            </tr>
                        </ItemTemplate>
                    </asp:ListView>
                </div>
            </ContentTemplate>
            <Triggers>
    <asp:AsyncPostBackTrigger ControlID="Records" EventName="ItemCommand" />
</Triggers>
        </asp:UpdatePanel>

        <!-- Delete Modal -->
        <div class="modal fade" id="deleteModal" tabindex="-1" aria-labelledby="modalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header bg-danger text-white">
                        <h5 class="modal-title" id="modalLabel">Confirm Delete</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        Are you sure you want to delete this item?
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="button" class="btn btn-danger" onclick="performDelete()">Delete</button>
                    </div>
                </div>
            </div>
        </div>

     <script type="text/javascript">
         var $realDeleteButton = null;
         var skipConfirm = false;

         function openConfirmModal(btn) {
             if (skipConfirm) {
                 // Allow postback
                 skipConfirm = false;
                 return true;
             }
             $realDeleteButton = $(btn);
             $('#deleteModal').modal('show');
             return false;
         }

         function performDelete() {
             $('#deleteModal').modal('hide');
             setTimeout(function () {
                 if ($realDeleteButton && $realDeleteButton.length) {
                     skipConfirm = true;
                     $realDeleteButton[0].click(); // Now OnClientClick will return true, allowing postback
                 }
             }, 200);
         }


     </script>

   
</asp:Content>
-----------------------------------------
-----------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication5
{
    public partial class WebForm2 : System.Web.UI.Page
    {
        public class Record
        {
            public string MPGROUP { get; set; }
            public string MPFIELD { get; set; }
            public string MPTRAN { get; set; }
            public string MPITVN { get; set; }
        }

        private static List<Record> _data = new List<Record>
    {
        // GroupA
        new Record { MPGROUP = "GroupA", MPFIELD = "Field1", MPTRAN = "Tran1", MPITVN = "Investor1" },
        new Record { MPGROUP = "GroupA", MPFIELD = "Field2", MPTRAN = "Tran2", MPITVN = "Investor2" },
        new Record { MPGROUP = "GroupA", MPFIELD = "Field3", MPTRAN = "Tran3", MPITVN = "Investor3" },
        new Record { MPGROUP = "GroupA", MPFIELD = "Field4", MPTRAN = "Tran4", MPITVN = "Investor4" },

        // GroupB
        new Record { MPGROUP = "GroupB", MPFIELD = "Field5", MPTRAN = "Tran5", MPITVN = "Investor5" },
        new Record { MPGROUP = "GroupB", MPFIELD = "Field6", MPTRAN = "Tran6", MPITVN = "Investor6" },
        new Record { MPGROUP = "GroupB", MPFIELD = "Field7", MPTRAN = "Tran7", MPITVN = "Investor7" },
        new Record { MPGROUP = "GroupB", MPFIELD = "Field8", MPTRAN = "Tran8", MPITVN = "Investor8" },

        // GroupC
        new Record { MPGROUP = "GroupC", MPFIELD = "Field9",  MPTRAN = "Tran9",  MPITVN = "Investor9" },
        new Record { MPGROUP = "GroupC", MPFIELD = "Field10", MPTRAN = "Tran10", MPITVN = "Investor10" },
        new Record { MPGROUP = "GroupC", MPFIELD = "Field11", MPTRAN = "Tran11", MPITVN = "Investor11" },
        new Record { MPGROUP = "GroupC", MPFIELD = "Field12", MPTRAN = "Tran12", MPITVN = "Investor12" }
    };

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                drpGroup.DataSource = new List<string> { "All", "GroupA", "GroupB", "GroupC" };
                drpGroup.DataBind();
                drpGroup.SelectedValue = "All";
                BindListView(_data);
            }
        }

        protected void btnGetRecords_Click(object sender, EventArgs e)
        {
            string selectedGroup = drpGroup.SelectedValue;
            List<Record> filtered = selectedGroup == "All"
                ? _data
                : _data.FindAll(r => r.MPGROUP == selectedGroup);
            BindListView(filtered);
        }

        private void BindListView(List<Record> records)
        {
            Records.DataSource = records;
            Records.DataBind();
        }

        protected void Records_ItemDeleting(object sender, ListViewDeleteEventArgs e)
        {
            // Prevent default delete behavior, since you handle it in ItemCommand
            e.Cancel = true;
        }

        protected void Records_ItemCommand(object sender, ListViewCommandEventArgs e)
        {
            try
            {
                if (e.CommandName == "MyDelete")
                {
                    string[] keys = e.CommandArgument.ToString().Split(',');
                    if (keys.Length == 4)
                    {
                        string group = keys[0];
                        string field = keys[1];
                        string tran = keys[2];
                        string itvn = keys[3];

                        _data.RemoveAll(r => r.MPGROUP == group && r.MPFIELD == field && r.MPTRAN == tran && r.MPITVN == itvn);
                    }

                    btnGetRecords_Click(null, null); // Refresh list with filter
                    updpMain.Update(); // Update the UI
                }
            }
            catch (Exception ex)
            {

                throw;
            }


        }

        protected void Records_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

===================================================

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" />
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
