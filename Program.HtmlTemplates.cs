// SPDX-License-Identifier: GPL-3.0-only

namespace SKAR_specs
{
    partial class Program
    {
        private static class Templates
        {
            public const string HtmlHeader = @"<!DOCTYPE html><html>
<head>
<meta charset=""utf-8"">
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
<title>SKAR_specs report</title>
<style type=""text/css"">
.tg  { border-collapse:collapse; border-color:#9ABAD9; border-spacing:0; margin:0px auto; }
.tg td{
    background-color:#EBF5FF; border-color:#9ABAD9; border-style:solid; border-width:1px; color:#444;
    font-family:Arial, sans-serif; font-size:14px; overflow:hidden; padding:10px 5px; word-break:normal;
}
.tg th{
    background-color:#409cff; border-color:#9ABAD9; border-style:solid; border-width:1px; color:#fff;
    font-family:Arial, sans-serif; font-size:14px; font-weight:normal; overflow:hidden; padding:10px 5px; word-break:normal;
}
.tg .tg-0lax{ text-align:left; vertical-align:top }
@media screen and (max-width: 767px) { .tg { width: auto !important; }.tg col { width: auto !important; }.tg-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; margin: auto 0px; } }

/* Sticky navigation styles */
.sticky-nav {
    position: sticky;
    top: 0;
    z-index: 100;
}

.sticky-nav th {
    background-color: #409cff;
}

.sticky-nav td {
    background-color: #EBF5FF;
}

/* Add a small shadow to make the sticky nav more visible when scrolling */
.sticky-nav-shadow {
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

/* Smooth scroll behavior for clicking on navigation links */
html {
    scroll-behavior: smooth;
}

/* Data cell specific styling */
.data-label { font-weight: bold; }
.data-value { }
.tg .data-value table {
    width: 100%;
    font-size: 12px;
    border-collapse: collapse;
    border-spacing: 0;
    margin: 0;
}

.tg .data-value table th,
.tg .data-value table td {
    font-family: Arial, sans-serif;
    font-size: 12px;
    line-height: 1.25;
    padding: 3px 5px;
    border: 1px solid #9ABAD9;
    word-break: normal;
}

.tg .data-value table th {
    background-color: #dbeafe;
    color: #000;
    font-weight: bold;
}

.tg .data-value table td {
    background-color: #fff;
    color: #444;
}
.spec-section { margin-top: 20px; }

@page {
    size: auto;
    margin: 12mm;
}

@media print {
    html {
        scroll-behavior: auto;
    }

    body {
        margin: 0;
        color: #000;
        background: #fff;
        font-size: 12px;
        line-height: 1.35;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }

    .tg-wrap {
        overflow: visible !important;
    }

    .tg {
        width: 100%;
        border-collapse: collapse;
        border-spacing: 0;
    }

    .tg td {
        background: #fff !important;
        color: #000;
    }

    .tg th {
        color: #000;
        background-color: #dbeafe !important;
    }

    .tg .data-value table {
        width: 100%;
        font-size: 12px;
        border-collapse: collapse;
        border-spacing: 0;
    }

    .tg .data-value table th,
    .tg .data-value table td {
        font-size: 12px;
        line-height: 1.25;
        padding: 3px 5px;
        border: 1px solid #9ABAD9;
        background: #fff !important;
        color: #000;
    }

    .tg .data-value table th {
        background-color: #dbeafe !important;
        font-weight: bold;
    }

    .sticky-nav {
        position: static;
    }

    .sticky-nav-shadow {
        box-shadow: none;
    }

    thead {
        display: table-header-group;
    }

    tfoot {
        display: table-footer-group;
    }

    tr,
    td,
    th {
        page-break-inside: avoid;
        break-inside: avoid;
    }

    h1, h2, h3, h4, header {
        page-break-after: avoid;
        break-after: avoid;
    }

    a,
    a:visited {
        color: inherit;
        text-decoration: none;
    }
}
</style>
</head>
<body>
<div class=""tg-wrap"" id=""system-specs-report""><table class=""tg"" role=""presentation""><thead>";

            public const string HtmlFooter = "</tbody></table></div></body></html>";

            public const string NavigationSection = @"
<tr class=""sticky-nav sticky-nav-shadow"">
<th class=""tg-0lax"" colspan=""2"" data-report-header=""true"">SKAR_specs | <span data-generated-date=""{0}"">{0}</span></th>
</tr>
</thead>
<tbody>
<tr class=""sticky-nav sticky-nav-shadow"">
<td class=""tg-0lax"" colspan=""2"" data-navigation=""main"">
  <nav role=""navigation"">
    [<a href=""#SUMMARY"">SUMMARY</a>]
    [<a href=""#Motherboard"">Motherboard</a>]
    [<a href=""#MEMORY"">MEMORY</a>]
    [<a href=""#HARDDRIVES"">HARD DRIVES</a>]
    [<a href=""#DISPLAY"">DISPLAY</a>]
    [<a href=""#Network"">NETWORK</a>]
    [<a href=""#NetworkShares"">Network Shares</a>]
    [<a href=""#Networkmappeddrives"">Network mapped drives</a>]
    [<a href=""#PRINTERS"">PRINTERS</a>]
    [<a href=""#SOFTWARE"">SOFTWARE</a>]
    [<a href=""#LICENSES"">LICENSES</a>]
  </nav>
</td>
</tr>";

            public const string SectionStart = @"
<tr class=""spec-section"" data-section-content=""{0}"">
<th class=""tg-0lax"" colspan=""2"" id=""{0}"" data-section=""{0}""><header>{1}</header></th>
</tr>";

            public const string TableRow = @"<tr>
<td class=""tg-0lax data-label"" data-key=""{0}"">{0}</td>
<td class=""tg-0lax data-value"" data-value=""{0}"">{1}</td>
</tr>";
        }
    }
}
