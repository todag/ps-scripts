#MIT License

#Copyright (c) 2017 https://github.com/todag

#Permission is hereby granted, free of charge, to any person obtaining a copy
#of this software and associated documentation files (the "Software"), to deal
#in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:
#The above copyright notice and this permission notice shall be included in all
#copies or substantial portions of the Software.

#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#SOFTWARE.

###############################################################################
#.SYNOPSIS
#
#   A simple script to display information on the users desktop.
#   The script displays a transparent WPF window that is set to a position
#   just above the desktop icons. Start it from the included .vbs script to
#   hide the Powershell console window.
#   It also includes a function to retrieve Active Directory attribute values
#   and display these on the desktop.
#
#.NOTES
#
#   If running on Powershell 2.0, it must be launched with -STA flag.
#
#   Version:        1.1
#
#   Author:         <https://github.com/todag>
#
#   Creation Date:  <2017-11-14>
#
#   1.0 Purpose/Change: Initial script development
#   1.1 Purpose/Change: Added SizeChanged event handler
#
###############################################################################

Set-StrictMode -Version 2.0
#
# Win32 methods to set window to backmost position, hide it from alt tab and hide it from mouse interaction
#
$Win32 = @"
    public static void SetupWindow(IntPtr windowHandle)
    {
        // Set window to backmost position (just above icons)
        SetWindowPos(windowHandle, new IntPtr(1), 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0010);

        // Hide from Alt Tab
        SetWindowLong(windowHandle, -20, 0x00000080);

        // Hide window from mouse interaction
        int extendedStyle = GetWindowLong(windowHandle, -20);
        SetWindowLong(windowHandle, -20, extendedStyle | 0x00000020);
    }

    [DllImport("user32.dll")]
    static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd, int index);
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);
"@

 [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window"
    Title="PS Desktop Info" Height="350" Width="525" WindowStyle="None" BorderBrush="Transparent" IsHitTestVisible="False" Background="{x:Null}" AllowsTransparency="True" WindowState="Maximized" ShowInTaskbar="False" AllowDrop="False">

    <Window.Resources>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
    </Window.Resources>

    <Grid HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,10,15,0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- This is a static TextBlock -->
        <TextBlock Grid.Row="0" Grid.ColumnSpan="2" Text="PC-Information"/>

        <!-- This Grid contains dynamically generated content -->
        <Grid Grid.Row="1" Name="grid">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
        </Grid>

        <!-- This is a static Grid with TextBlocks -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Grid.Column="0"/>
            <TextBlock Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Text="For IT-support, contact helpdesk:"/>
            <TextBlock Grid.Row="2" Grid.Column="0" Text="E-mail: "/>
            <TextBlock Grid.Row="2" Grid.Column="1" Text="helpdesk@company.com"/>
            <TextBlock Grid.Row="3" Grid.Column="0" Text="Web: "/>
            <TextBlock Grid.Row="3" Grid.Column="1" Text="helpdesk.company.com"/>
            <TextBlock Grid.Row="4" Grid.Column="0" Text="Phone: "/>
            <TextBlock Grid.Row="4" Grid.Column="1" Text="0123-456 78"/>
        </Grid>
    </Grid>
</Window>
'@

#
# Add required types and assemblies
#
Add-Type -MemberDefinition $Win32 -Name Win32 -Namespace System
Add-Type –AssemblyName PresentationFramework

#
# Setup Window
#
$Window = @{}
$Window.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
    $Window.$($_.Name) = $Window.Window.FindName($_.Name)
}

#
# Add Loaded event handler
# This will set the Window position, hide it from Alt Tab and mouse interaction.
#
$Window.Window.add_Loaded({
    $handle =  (New-Object System.Windows.Interop.WindowInteropHelper($Window.Window)).Handle
    [Win32]::SetupWindow($handle);
})

#
# Add SizeChanged event handler. Will be called if size changes
# and reset window to maximized.
#
$Window.Window.add_SizeChanged({
    $Window.Window.WindowState = "Maximized"
})

#
# Open directory entry so we can get attribute values
#
$adEntry = $null
try
{
    $adSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $adSearcher.Filter = "(&(objectCategory=Computer)(SamAccountname=$($env:computerName)`$))"
    $adEntry =  $adSearcher.FindOne().GetDirectoryEntry()
    $adSearcher.Dispose()
}
catch{}

#
# If directoryEntry has been opened successfully, this retrieves attribute values.
#
function Get-AttributeValue([string]$attribute)
{
    if($adEntry -ne $null)
    {
        return $adEntry.Properties[$attribute].ToString()
    }
    else
    {
        return "<Unable to load data>"
    }
}

#
# Setup array containing all data
#
$dataArray = @(
    ("Datornamn: ",     $env:COMPUTERNAME),
    ("Manufacturer: ",  (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Manufacturer),
    ("Model: ",         (Get-WmiObject Win32_ComputerSystem -ErrorAction SilentlyContinue).Model),
    ("Serial#: ",       (Get-WMIObject Win32_BIOS -ErrorAction SilentlyContinue).SerialNumber),

    #Remove these if not running on domain joined computer
    ("Companyg ",       (Get-AttributeValue("company"))),
    ("Division: ",      (Get-AttributeValue("division"))),
    ("Department: ",    (Get-AttributeValue("department"))),
    ("Dept. Number: ",  (Get-AttributeValue("departmentNumber"))),
    ("Location: ",      (Get-AttributeValue("location")))
)

#
# Close directory entry
#
if($adEntry -ne $null)
{
    $adEntry.Dispose()
}

#
# Generate TextBlocks from data in $dataArray and add to Grid
#
[int]$row = 0
foreach($data in $dataArray)
{
    $rowDefinition = New-Object System.Windows.Controls.RowDefinition
    $txtKey = New-Object System.Windows.Controls.TextBlock
    $txtKey.SetValue([Windows.Controls.Grid]::ColumnProperty, 0);
    $txtKey.SetValue([Windows.Controls.Grid]::RowProperty, $row);
    $txtKey.Text = $data[0]

    $txtValue = New-Object System.Windows.Controls.TextBlock
    $txtValue.SetValue([Windows.Controls.Grid]::ColumnProperty, 1);
    $txtValue.SetValue([Windows.Controls.Grid]::RowProperty, $row);
    $txtValue.Text = $data[1]

    $Window.grid.RowDefinitions.Add($rowDefinition) | Out-Null
    $Window.grid.Children.Add($txtKey) | Out-Null
    $Window.grid.Children.Add($txtValue) | Out-Null
    $row++
}

$Window.Window.ShowDialog() | out-null
