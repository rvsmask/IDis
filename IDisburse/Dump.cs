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
