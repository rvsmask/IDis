string userName = HttpContext.Current.User.Identity.Name;

protected void Page_Load(object sender, EventArgs e)
{
    string userName = Page.User.Identity.Name; // e.g., DOMAIN\username
    Label1.Text = "Logged in as: " + userName;

    if (Page.User.IsInRole("DOMAIN\\Admins"))
    {
        Label2.Text = "Role: Admin";
    }
    else if (Page.User.IsInRole("DOMAIN\\Users"))
    {
        Label2.Text = "Role: User";
    }
    else
    {
        Label2.Text = "Role: Unknown";
    }
}
