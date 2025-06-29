<!-- Bootstrap 4 or 5 (choose one) -->
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" />
<script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

  <ItemTemplate>
                <asp:Button ID="btnDelete" runat="server" Text="Delete" CommandName="Delete"
                    CommandArgument='<%# Eval("ID") %>'
                    CssClass="btn btn-danger btn-sm"
                    OnClientClick="openConfirmModal(this); return false;" />
            </ItemTemplate>

                    -------------

                    <!-- Delete Confirmation Modal -->
<div class="modal fade" id="deleteModal" tabindex="-1" role="dialog" aria-labelledby="modalLabel" aria-hidden="true">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header bg-danger text-white">
        <h5 class="modal-title" id="modalLabel">Confirm Delete</h5>
        <button type="button" class="close text-white" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        Are you sure you want to delete this item?
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancel</button>
        <button type="button" class="btn btn-danger" onclick="performDelete()">Delete</button>
      </div>
    </div>
  </div>
</div>

--------------

<script type="text/javascript">
    var clickedButton = null;

    function openConfirmModal(btn) {
        clickedButton = btn; // store button for later
        $('#deleteModal').modal('show');
    }

    function performDelete() {
        $('#deleteModal').modal('hide');
        __doPostBack(clickedButton.name, '');
    }
</script>

-----------------

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Delete Confirmation Modal Demo</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server" />

        <div class="container mt-5">
            <asp:ListView ID="ListView1" runat="server" DataKeyNames="ID">
                <ItemTemplate>
                    <div class="d-flex justify-content-between align-items-center border-bottom py-2">
                        <span><%# Eval("Name") %></span>
                        <asp:LinkButton ID="LinkButtonDelete" runat="server" CssClass="btn btn-danger btn-sm delete-btn"
                            CommandName="Delete" CommandArgument='<%# Eval("ID") %>'>
                            Delete
                        </asp:LinkButton>
                    </div>
                </ItemTemplate>
            </asp:ListView>
        </div>

        <!-- Delete Confirmation Modal -->
        <div class="modal fade" id="deleteModal" tabindex="-1" aria-labelledby="deleteModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="deleteModalLabel">Confirm Delete</h5>
                    </div>
                    <div class="modal-body">
                        Are you sure you want to delete this item?
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="button" class="btn btn-danger" id="confirmDeleteBtn">Delete</button>
                    </div>
                </div>
            </div>
        </div>
    </form>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script type="text/javascript">
        document.addEventListener('DOMContentLoaded', function () {
            var targetBtn = null;

            document.querySelectorAll('.delete-btn').forEach(function (btn) {
                btn.addEventListener('click', function (e) {
                    e.preventDefault();
                    targetBtn = btn;
                    var modal = new bootstrap.Modal(document.getElementById('deleteModal'));
                    modal.show();
                });
            });

            document.getElementById('confirmDeleteBtn').addEventListener('click', function () {
                if (targetBtn) {
                    __doPostBack(targetBtn.name, '');
                }
            });
        });
    </script>
</body>
</html>

