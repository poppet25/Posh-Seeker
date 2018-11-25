$sb = @'
    <script type="text/javascript">
    $(document).ready(function(){
        $('table').each(function(){
            // Grab the contents of the first TR element and save them to a variable
            var tHead = $(this).find('tr:first').html();
            // Remove the first COLGROUP element 
            $(this).find('colgroup').remove(); 
            // Remove the first TR element 
            $(this).find('tr:first').remove();
            // Add a new THEAD element before the TBODY element, with the contents of the first TR element which we saved earlier. 
            $(this).find('tbody').before('<thead>' + tHead + '</thead>'); });
            // Apply the DataTables jScript to all tables on the page 
            $('table').dataTable( {
            // Put your datatable options here 
        } ); });
    </script>
'@
$css = @(
    '<link rel="stylesheet" type="text/css" href="js/DataTables/datatables.min.css">',
    '<link rel="stylesheet" type="text/css" href="js/bootstrap-4.1.3-dist/css/bootstrap.min.css">'
    )
$js = @(
    '<script src="js/jquery-3.3.1.min.js" ></script>',
    '<script src="js/DataTables/datatables.min.js" ></script>',
    '<script src="js/bootstrap-4.1.3-dist/js/bootstrap.min.js" ></script>',
    $sb
    )
$body = @'
        <div class="card card-body bg-secondary">
            <h1>PII Searcher</h1>
        </div>
'@

$meta = @{
    'Content-Type' = "text/html"
}

Start-Job -ScriptBlock {
    0..10000 | ForEach-Object{ New-Object psobject -Property @{ Value = $_; Name="Name$_" }} | `
        ConvertTo-Html -Meta $meta  -Head $css -PostContent $js -PreContent $body -Charset utf8 | `
        Out-File -FilePath "file.htm"
}