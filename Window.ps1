Add-Type -AssemblyName PresentationFramework
  
 
$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   SizeToContent="WidthAndHeight"
   Title="Get Out!"
   Topmost="True">
      <TextBlock
         Margin="50"
         HorizontalAlignment="Center"
         VerticalAlignment="Center"
         FontFamily="Stencil"
         FontSize="80"
         FontWeight="Bold"
         Foreground="Red">
         Fire Alarm!
      </TextBlock>
</Window>
'@
 
$reader = [System.XML.XMLReader]::Create([System.IO.StringReader]$XAML)
$window = [System.Windows.Markup.XAMLReader]::Load($reader)
  
$window.ShowDialog() 
