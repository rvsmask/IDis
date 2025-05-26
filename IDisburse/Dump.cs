<asp:HiddenField ID="hfIsVisible" runat="server" />

<div id="myDiv" style="display: none;">
    This is a sample div.
</div>

<script type="text/javascript">
    window.onload = function () {
        var isVisible = document.getElementById('<%= hfIsVisible.ClientID %>').value;

        var div = document.getElementById('myDiv');
        if (isVisible === "true") {
            div.style.display = "block";
        } else {
            div.style.display = "none";
        }
    };
</script>

protected void Page_Load(object sender, EventArgs e)
{
    // Your server-side condition
    bool isVisible = SomeCondition(); // e.g., check session, database, etc.

    hfIsVisible.Value = isVisible.ToString().ToLower(); // "true" or "false"
}

protected void Page_Load(object sender, EventArgs e)
{
    bool isVisible = SomeCondition();

    string script = $"<script>window.onload = function() {{ " +
                    $"document.getElementById('myDiv').style.display = '{(isVisible ? "block" : "none")}'; }};</script>";

    ClientScript.RegisterStartupScript(this.GetType(), "ShowHideDiv", script);
}
//-------------23--------------

    protected void btnFilter_Click(object sender, EventArgs e)
{
    lblValidationError.Visible = false;

    string fromDateText = txtFromDate.Text.Trim();
    string toDateText = txtToDate.Text.Trim();

    DateTime fromDate, toDate;

    bool fromValid = DateTime.TryParse(fromDateText, out fromDate);
    bool toValid = DateTime.TryParse(toDateText, out toDate);

    if (!string.IsNullOrEmpty(fromDateText) && !fromValid)
    {
        lblValidationError.Text = "Please enter a valid 'From Date'.";
        lblValidationError.Visible = true;
        return;
    }

    if (!string.IsNullOrEmpty(toDateText) && !toValid)
    {
        lblValidationError.Text = "Please enter a valid 'To Date'.";
        lblValidationError.Visible = true;
        return;
    }

    if (fromValid && toValid && fromDate > toDate)
    {
        lblValidationError.Text = "'From Date' cannot be after 'To Date'.";
        lblValidationError.Visible = true;
        return;
    }

    BindGrid(); // Only bind if validation passes
}
