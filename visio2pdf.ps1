#
# Given one or more visio documents,
# Crops them with 0 margin and exports as PDF.
# The original documents are left unaltered, PDFs are overwritten.
#

Param (
  [parameter(ValueFromPipeline=$true)][string[]] $InputPaths = "*.vsdx"
)


function autocrop {
    param( $document )

    $document.TopMargin("mm") = 0
    $document.BottomMargin("mm") = 0
    $document.LeftMargin("mm") = 0
    $document.RightMargin("mm") = 0

    Foreach ($page in $document.Pages) {
        $page.ResizeToFitContents()
    }    
    
}

function export_pdf {
    param ($document,
           [string] $outfile )
           
    $document.ExportAsFixedFormat(1, # PDF
                                  $outfile,
                                  1, # Print quality
                                  0  # All pages
                                 )
    

}

function crop_and_export {
    param ( $visio,
            [string] $infile,
            [string] $outfile )
    echo "Converting $infile to $outfile"
    $document = $visio.Documents.Add($infile)
    
    autocrop $document
    export_pdf $document $outfile

    # Automatically respond "No" instead of showing the "Save" dialog
    # Set to 6 for "Yes"
    $visio.AlertResponse = 7;
    $document.Close()
    $visio.AlertResponse = 0;
    
}

# Take care of wildcards in the input
$InputFiles = @()
Foreach( $path in $InputPaths) {
    $globbed = Resolve-Path $path
    $InputFiles += $globbed
}



$visio = New-Object -ComObject Visio.InvisibleApp
#$visio = New-Object -ComObject Visio.Application


ForEach( $infile in $InputFiles ) {
    $outfile = $infile -replace "\..*$",""
    $outfile = $outfile + ".pdf"
    
    crop_and_export $visio $infile $outfile

}

#Read-Host "Hit RETURN to quit"

$visio.Quit()