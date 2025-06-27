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
