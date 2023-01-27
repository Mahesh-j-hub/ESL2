<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm.aspx.cs" Inherits="ESL_Chart.WebForm" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="ChartCSS.css" rel="stylesheet" />
    <link rel="icon" type="image/x-icon" href="https://img.icons8.com/plasticine/2x/analytics.png" />
    <link href="Content/bootstrap.css" rel="stylesheet" />
    <script src="Scripts/jquery-3.6.3.js"></script>
    <script src="Scripts/jquery-3.6.3.min.js"></script>
    <link href="Scripts/fontawesome-free-6.2.0-web/css/fontawesome.css" rel="stylesheet" />
    <link href="Scripts/fontawesome-free-6.2.0-web/css/all.css" rel="stylesheet" />
    <title>ESL</title>
    <style>
        @keyframes moveDown {
            0% {
                opacity: 0;
                transform: translateY(-0.6rem);
            }

            100% {
                transform: translateY(0.6rem);
            }
        }

        @keyframes moveInLeft {
            0% {
                opacity: 0;
                transform: translateY(-1rem);
            }

            80% {
                transform: translateY(3rem);
            }

            100% {
                opacity: 1;
                transform: translate(0);
            }
        }

        #goDown:hover #icondown {
            text-decoration: none;
            animation-name: moveDown !important;
            animation-duration: 1s !important;
            box-shadow: 0 0 10px rgba(0,0,0,0.5);
        }

        .dropdown1 {
            float: left;
            overflow: hidden;
        }

            .dropdown1 .dropbtn1 {
                font-size: 16px;
                border: none;
                outline: none;
                color: white;
                padding: 14px 16px;
                background-color: inherit;
                font-family: inherit;
                margin: 0;
            }

            .navbar a:hover, .dropdown1:hover .dropbtn1 {
                background-color: grey;
            }

        .dropdown-content1 {
            /*display: block;*/
            position: absolute;
            background-color: Black;
            min-width: 160px;
            box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
            z-index: 1;
            overflow: auto;
        }

            .dropdown-content1 a {
                float: none;
                background-color: black;
                color: white;
                padding: 12px 16px;
                text-decoration: none;
                display: grid;
                text-align: left;
            }

                .dropdown-content1 a:hover {
                    background-color: red;
                    color: black;
                    display: block;
                }

        .dropdown1:hover .dropdown-content1 {
            display: block;
        }
    </style>
    <script>
        $(document).ready(function () {
            $('#loadingImage').hide();
        });
    </script>
</head>
<body>
    <asp:Image CssClass="loaderImg" ID="loadingImage" runat="server" ImageUrl="loader.gif" />
    <form id="form1" runat="server">
        <%-- Navigation Bar --%>
        <nav>
            <div class="heading">
                <i class="fa fa-line-chart " style="font-size: 36px"></i>&nbsp; Employees Skill Ladder
            <h6 style="font-size: 12px">ESL</h6>
            </div>
        </nav>
        <ul>
            <asp:LinkButton Style="float: right; background-color: #333; color: white; border: none; margin-top: 15px; margin-right: 12px; margin-left: 10px" ID="btnlogout" OnClick="Logout" runat="server"><i class="fa-solid fa-right-from-bracket" aria-hidden="true"></i> Logout</asp:LinkButton>
            <li style="float: right"><a href="/Home/Edit/<%=Session["Identity"] %>">Welcome <%=Session["User"] %></a></li>
            <li class="dropdown1"><a class="dropbtn1">File Manager</a>
                <div class="dropdown-content1" hidden="hidden">
                    <a href="/Home1/Upload">Upload</a>
                    <a href="/Home1/Log">Log</a>
                    <a href="/Home1/View">View</a>
                </div>
            </li>
            <li><a href="/Home1/Tracker">Tracker</a></li>
            <li><a href="/Home1/Summary">Summary</a></li>
            <li><a href="#">Chart</a></li>
            <li class="dropdown1">
                <a class="dropbtn1">Admin</a>
                <div class="dropdown-content1" hidden="hidden">
                    <a href="/Home/Edit/<%=Session["Identity"] %>">User Management</a>
                    <a href="Roll/Index">Role Master</a>
                    <a href="RollMapping/Index">Role Mapping</a>
                </div>
            </li>
        </ul>

        <div>
            <asp:DropDownList ID="DDTest" runat="server">
                <asp:ListItem Selected="True">Select Test Type</asp:ListItem>
                <asp:ListItem>FCA Programming Language</asp:ListItem>
                <asp:ListItem>FCA Software Engineering</asp:ListItem>
                <asp:ListItem>SME Assessment</asp:ListItem>
            </asp:DropDownList>

            <asp:DropDownList ID="DDDimension" runat="server">
                <asp:ListItem Selected="True">Select Chart Dimension</asp:ListItem>
                <asp:ListItem>2 Dimensional Chart</asp:ListItem>
                <asp:ListItem>3 Dimensional Chart</asp:ListItem>
            </asp:DropDownList>

            <asp:DropDownList ID="DDBatch" runat="server">
                <asp:ListItem Selected="True">Select Batches</asp:ListItem>
            </asp:DropDownList>

            <asp:DropDownList ID="DDChartType" runat="server">
                <asp:ListItem Selected="True">Select Chart type</asp:ListItem>
            </asp:DropDownList>

            <asp:DropDownList OnSelectedIndexChanged="DDAttemptsSelectedIndexChanged" AutoPostBack="true" Visible="false" ID="DDAttempts" runat="server">
                <asp:ListItem Selected="True">Attempt 1</asp:ListItem>
                <asp:ListItem>Attempt 2</asp:ListItem>
                <asp:ListItem>Attempt 3</asp:ListItem>
            </asp:DropDownList>
            <br />
            <asp:Button ID="Button1" CssClass="btn btn-primary" runat="server" OnClick="GenerateGraphEvent" Text="Generate Chart" />
            <asp:Button ID="Button4" CssClass="btn btn-success" OnClick="GoBackToMainCharts" Text="Go Back to Main Charts" runat="server" Visible="false" />
            <asp:Button ID="Button2" CssClass="btn btn-primary" runat="server" OnClick="DetailEvent" Text="Detail" />
            <asp:Button ID="Button3" CssClass="btn btn-primary" OnClientClick="Save()" Text="Save" OnClick="Save" runat="server" />
            <a visible="false" id="goDown" style="float: right; text-decoration: none; margin-right: 200px; margin-top: 20px; animation-name: moveInLeft; animation-duration: 1.2s;" href="#Chart3" runat="server"><i id="icondown" class="fa fa-arrow-circle-down" aria-hidden="true"></i>&nbsp;More Info</a>
            <br />
            <asp:Chart ID="Chart1" runat="server" Width="500px" Height="350px">
                <Series>
                    <asp:Series Name="Series1">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1">
                    </asp:ChartArea>
                </ChartAreas>
            </asp:Chart>

            <asp:Chart ID="Chart2" runat="server" Width="500px" Height="350px">
                <Series>
                    <asp:Series Name="Series2">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea2">
                    </asp:ChartArea>
                </ChartAreas>
            </asp:Chart>

            <asp:Chart ID="Chart3" Visible="false" runat="server" Width="800px" Height="350px">
                <Series>
                    <asp:Series Name="Attempt 1">
                    </asp:Series>
                    <asp:Series Name="Attempt 2">
                    </asp:Series>
                    <asp:Series Name="Attempt 3">
                    </asp:Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="ChartArea1">
                    </asp:ChartArea>
                </ChartAreas>
            </asp:Chart>
            &nbsp;
            <br />
            <br />
            <div style="text-align: left">
                <div id="Grid1">
                    <asp:Label Style="font-weight: bold" ID="label1" Text="" runat="server" />
                    <asp:GridView ID="GridView1" runat="server">
                    </asp:GridView>
                </div>
                <div id="Grid2">
                    <asp:Label Style="font-weight: bold" ID="label2" Text="" runat="server" />
                    <asp:GridView ID="GridView2" runat="server">
                    </asp:GridView>
                </div>
                <div id="Grid3">
                    <asp:Label Style="font-weight: bold" ID="label3" Text="" runat="server" />
                    <asp:GridView ID="GridView3" runat="server">
                    </asp:GridView>
                </div>
            </div>
        </div>
        <div style="width: 50%; margin-left: 10px">
            <p id="para" visible="false" runat="server" style="font-size: 10px; font-style: italic;">
                <i title="Info" class="fa fa-info-circle" aria-hidden="true"></i>
                <span style="opacity: 0.8;">Table describes total number of people cleard Test in <b>particular TCB based on Attempts.</b></span>
            </p>
        </div>
    </form>

    <footer class="page-footer font-small unique-color-dark text-center" id="footer" style="background-color: #48a0d9;">
        <div style="background-color: #48a0d9;">
            <div class="container">
                <div class="row py-4 d-flex align-items-center">
                    <!--Grid column-->
                    <div class="col-md-6 col-lg-5 text-center text-md-center mb-4 mb-md-0" style="color: white; margin-left: 380px">
                        © <%= DateTime.Now.ToString("yyyy") %> Copyright Sasken all rights reserved 
                    </div>
                    <!--Grid column-->
                </div>
            </div>
        </div>
    </footer>
</body>
</html>
