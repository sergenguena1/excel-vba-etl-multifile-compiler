' =============================================================================
' PROJECT  : Multi-File ETL Automation — Excel VBA
' AUTHOR   : Serge NGUENA
' CONTEXT  : Developed at Danone Sub-Saharan Africa to process 70+ distributor
'            VMI (Vendor-Managed Inventory) files monthly and compile them into
'            a single structured database ready for Power BI ingestion.
' RESULT   : Full processing of 70+ files completed in under 4 minutes,
'            replacing a manual process that previously took several hours.
' =============================================================================
'
' HOW IT WORKS
' ------------
' 1. Opens a master Template file containing the standardized formula logic
'    to be applied uniformly across all source files.
' 2. Deletes and recreates the output BDD (database) file from scratch
'    to ensure a clean, non-duplicated compilation on every run.
' 3. Loops through every .xlsx file in the target folder:
'      a. Opens the file
'      b. Pastes the template formula range (values & structure normalization)
'      c. Saves the processed file
'      d. Copies the cleaned data rows into the master BDD
'      e. Closes the file
' 4. Saves and closes the final BDD file and notifies the user.
'
' INPUT    : 70+ .xlsx files with identical column structure in a single folder
' OUTPUT   : One consolidated .xlsx database file (BDD), structured for BI tools
' =============================================================================

Option Explicit

Sub TraiterEtCompiler()

    ' -------------------------------------------------------------------------
    ' CONFIGURATION — Update these paths and filenames to match your environment
    ' -------------------------------------------------------------------------
    Dim cheminDossier As String
    cheminDossier = "C:\YourPath\SourceFiles\"                          ' Folder containing all source files

    Dim nomFichierTemplate As String
    nomFichierTemplate = "TEMPLATE_VMI_Structure.xlsm"                  ' Master template with formula logic

    Dim nomFichierBDD As String
    nomFichierBDD = "Compil_VMI_BDD_Output.xlsx"                        ' Output database filename

    ' -------------------------------------------------------------------------
    ' VARIABLE DECLARATIONS
    ' -------------------------------------------------------------------------
    Dim classeurTemplate As Workbook        ' Template workbook (formula logic)
    Dim classeurBDD As Workbook             ' Output BDD workbook
    Dim classeurCourant As Workbook         ' Current source file being processed
    Dim feuilleBDD As Worksheet             ' BDD sheet reference
    Dim derniereLigne As Long               ' Tracks last used row in BDD
    Dim fichier As String                   ' Current filename in loop
    Dim compteur As Integer                 ' File counter for progress tracking

    ' -------------------------------------------------------------------------
    ' STEP 1 — Open the Template file (contains the normalization formula range)
    ' -------------------------------------------------------------------------
    On Error Resume Next
    Set classeurTemplate = Workbooks(nomFichierTemplate)
    On Error GoTo 0

    If classeurTemplate Is Nothing Then
        Set classeurTemplate = Workbooks.Open(cheminDossier & nomFichierTemplate)
    End If

    ' -------------------------------------------------------------------------
    ' STEP 2 — Reset output BDD: close if open, delete if exists, recreate fresh
    ' This guarantees no duplicate data across runs.
    ' -------------------------------------------------------------------------
    On Error Resume Next
    Workbooks(nomFichierBDD).Close SaveChanges:=False
    On Error GoTo 0

    If Dir(cheminDossier & nomFichierBDD) <> "" Then
        Kill cheminDossier & nomFichierBDD
    End If

    Set classeurBDD = Workbooks.Add
    classeurBDD.SaveAs cheminDossier & nomFichierBDD, FileFormat:=xlOpenXMLWorkbook

    Set feuilleBDD = classeurBDD.Sheets(1)
    feuilleBDD.Name = "BDD"

    ' -------------------------------------------------------------------------
    ' STEP 3 — Loop through all .xlsx files in the source folder
    ' -------------------------------------------------------------------------
    fichier = Dir(cheminDossier & "*.xlsx")
    compteur = 0

    Do While fichier <> ""

        ' Skip the output BDD file if it is located in the same folder
        If fichier <> nomFichierBDD Then

            ' Open the current source file
            Set classeurCourant = Workbooks.Open(cheminDossier & fichier)

            ' --- NORMALIZE: Apply template formula range to current file ---
            ' Pastes the standardized structure from the template into the source file.
            ' This ensures all files share the same calculated columns before compilation.
            classeurTemplate.Sheets("Sales").Range("Q9:AM475").Copy _
                classeurCourant.Sheets("Sales").Range("Q9")
            Application.CutCopyMode = False

            ' Save the normalized file
            classeurCourant.Save

            ' --- COMPILE: Append cleaned data rows to the BDD ---
            ' Only values are pasted (no formulas) to keep the BDD lightweight and portable.
            derniereLigne = feuilleBDD.Cells(feuilleBDD.Rows.Count, "A").End(xlUp).Row
            classeurCourant.Sheets("Sales").Range("Q11:AH475").Copy
            feuilleBDD.Cells(derniereLigne + 1, 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False

            ' Close the processed file without re-saving (already saved above)
            classeurCourant.Close SaveChanges:=False

            compteur = compteur + 1

        End If

        ' Move to the next file
        fichier = Dir

    Loop

    ' -------------------------------------------------------------------------
    ' STEP 4 — Save and close the final BDD, notify user
    ' -------------------------------------------------------------------------
    classeurBDD.Close SaveChanges:=True

    MsgBox "Processing complete!" & vbNewLine & vbNewLine & _
           "Files processed : " & compteur & vbNewLine & _
           "Output file     : " & nomFichierBDD, vbInformation

End Sub
