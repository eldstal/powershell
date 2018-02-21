#
# Given one or more visio documents,
# Crops each page with 0 margin and exports as PDF.
# The original documents are left unaltered, PDFs are overwritten.
#

Param (
  [parameter(ValueFromPipeline=$true)][string[]] $InputPaths = "*.vsdx",
  [parameter(Mandatory=$false)][Switch] $Visible = $false
)


function zero_margin {
    param( $document )

    $document.TopMargin("mm") = 0
    $document.BottomMargin("mm") = 0
    $document.LeftMargin("mm") = 0
    $document.RightMargin("mm") = 0  
    
}

function export_pdf {
    param ($document,
           $page_number,
           [string] $outfile )
           
    $document.ExportAsFixedFormat(1, # PDF format
                                  $outfile,
                                  1, # Print quality
                                  1, # Print a page range, rather than all pages
                                  $page_number, # Starting page
                                  $page_number  # Ending page
                                 )
    

}

function crop_and_export {
    param ( $visio,
            [string] $infile,
            [string] $outdir )
    echo "Converting $infile and saving figures to $outdir"
    
    # Open the source file
    $document = $visio.Documents.Add($infile)
    
    zero_margin $document
    
    Foreach ($page in $document.Pages) {
    
        $filename = $page.Name -replace " ","_"
        $filename = $filename + ".pdf"
        $outfile = Join-Path -Path $outdir -ChildPath $filename
    
        $page.ResizeToFitContents()
        echo "  $outfile"
        export_pdf $document $page.Index $outfile
    }    

    # Automatically respond "No" instead of showing the "Save" dialog
    # Set AlertResponse to 6 for "Yes"
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


$class = "Visio.InvisibleApp"
if ($Visible) {
    $class = "Visio.Application"
}

# Launch a visio instance that we can control
$visio = New-Object -ComObject $class


ForEach( $infile in $InputFiles ) {
    $outfile = $infile -replace "\..*$",""
    $outdir = Split-Path -Path $infile -Parent
    
    crop_and_export $visio $infile $outdir

}

#Read-Host "Hit RETURN to quit"

$visio.Quit()