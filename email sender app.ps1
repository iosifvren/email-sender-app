Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Alex' application for the Global Reporting and Analysis Team!"
$form.Size = New-Object System.Drawing.Size(600, 500)

$labels = @("Select Report Template", "To", "CC", "BCC", "Subject", "Body", "Attachment")
$yPos = 20
foreach ($label in $labels) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $label
    $lbl.Location = New-Object System.Drawing.Point(10, $yPos)
    $lbl.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lbl)
    $yPos += 30
}

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(140, 20)
$comboBox.Size = New-Object System.Drawing.Size(400, 20)
$comboBox.Items.AddRange(@(
    "Report1"
    "Report2"
    "Report3", 
    
))
$form.Controls.Add($comboBox)

$fields = @("To", "CC", "BCC", "Subject", "Body", "Attachment")
$yPos = 50
foreach ($field in $fields) {
    $txtBox = New-Object System.Windows.Forms.TextBox
    $txtBox.Name = $field
    $txtBox.Location = New-Object System.Drawing.Point(140, $yPos)
    $txtBox.Size = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($txtBox)
    $yPos += 30
}

$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = "Copyright © 2024 Iosif-Alexandru Vrencean`nAll rights reserved"
$footerLabel.AutoSize = $true
$footerLabel.Location = New-Object System.Drawing.Point(10, 450)
$form.Controls.Add($footerLabel)

$templates = @{

    "Report1" = @{
        "To" = "iosifa286@gmail.com"
        "CC" = ""
        "BCC" = ""
        "Subject" = "Report1"
        "Body" = "<div style='font-size:9pt; font-family:Biennale'>Hi Both,<br><br>The Finance Report is uploaded in: <a href='https://atos365.sharepoint.com/:f:/r/sites/100005485/Shared%20Documents/02%20%20Leadership%20Team/Data%20Integrity%20and%20Reports/A%20FIN?csf=1&web=1&e=50n6gq'>A FIN REPORT</a><br><br></div><br><table cellpadding='0' cellspacing='0' border='0' width='100%' id='tablePreview' style='display: table; margin-bottom: 0px; transition: 300ms; font-size:9pt; font-family:Biennale;'><tbody><tr><td style='margin:0 0 15px 0;padding:0' id='politeSignature' colspan='2'><div id='polite1' style='margin:10px 0;font-family:Biennale;color:#000;font-size:9pt'>Kind regards,</div></td></tr><tr><td id='logoSignature' style='vertical-align: middle; padding-right: 20px; width: 150px;'><span><img width='150' src='https://atos.net/content/email-signature/euro-2024-atos.png' id='logo' alt='equensWorldline logo'></span></td><td width='100%' style='border-left: 2px solid rgb(128, 128, 128); max-width: 100%; padding-left: 25px; line-height: 1; padding-bottom: 0px;' id='textSignature'><span id='fullName1' style='font-family:Biennale;color:#0073e6;font-weight:bold;font-size:13px'>Iosif-Alexandru Vrencean</span><br><span id='jobtitle1' style='font-family:Biennale;color:#000;font-size:11px'>Global Reporting Analyst</span><br><span style='font-family:Biennale;color:#000;font-size:11px' id='address1'>Atos IT Solutions and Services EOOD – 2 Knyaginya Maria Louis\a Blvd Sofia </span><span style='font-family:Biennale;color:#000;font-size:11px' id='addressline2span'> – 1000 – Bulgaria</span><span style='font-family:Biennale;color:#000;font-size:11px' id='country1'> – Bulgaria</span><br><a href='https://atos.net/bg' target='_blank' style='font-family:Biennale;font-size:11px;text-decoration:underline;color:#0073e6'>atos.net/bg</a><br><span id='idWrapIcons' class='wrapIcons' style='font-size:22px;margin-top:6px'><span><a target='_blank' href='https://www.linkedin.com/in/iosif-alexandru-vrencean-3747211b8/'><img src='https://atos.net/content/email-signature/linkedin.png' width='15' height='15' alt='LinkedIn icon'></a></span>&nbsp;<span><a title='LinkedIn' href='https://twitter.com/atos' target='_blank'><img src='https://atos.net/content/email-signature/twitter.png' width='15' height='15' alt='Twitter icon'></a></span>&nbsp;<span><a href='https://www.facebook.com/Atos/' target='_blank'><img src='https://atos.net/content/email-signature/facebook.png' width='15' height='15' alt='Facebook icon'></a></span>&nbsp;<span><a href='https://www.instagram.com/atosinside/' target='_blank'><img src='https://atos.net/content/email-signature/instagram.png' width='15' height='15' alt='Instagram icon'></a></span>&nbsp;<span><a href='https://www.youtube.com/user/Atos' target='_blank'><img src='https://atos.net/content/email-signature/youtube.png' width='15' height='15' alt='Youtube icon'></a></span>&nbsp;<span><a href='https://www.xing.com/companies/atosorigin' target='_blank'><img src='https://atos.net/content/email-signature/xing.png' width='15' height='15' alt='Xing icon'></a></span>&nbsp;</span></td></tr></tbody></table>"
    }
    
    "Report2" = @{
        "To" = "iosifa286@gmail.com"
        "CC" = ""
        "BCC" = "t"
        "Subject" = "Report2"
        "Body" = "<div style='font-size:9pt; font-family:Biennale'>Hi Maria,<br><br>Please find attached the report.<br><br></div><br><table cellpadding='0' cellspacing='0' border='0' width='100%' id='tablePreview' style='display: table; margin-bottom: 0px; transition: 300ms; font-size:9pt; font-family:Biennale;'><tbody><tr><td style='margin:0 0 15px 0;padding:0' id='politeSignature' colspan='2'><div id='polite1' style='margin:10px 0;font-family:Biennale;color:#000;font-size:9pt'>Kind regards,</div></td></tr><tr><td id='logoSignature' style='vertical-align: middle; padding-right: 20px; width: 150px;'><span><img width='150' src='https://atos.net/content/email-signature/euro-2024-atos.png' id='logo' alt='equensWorldline logo'></span></td><td width='100%' style='border-left: 2px solid rgb(128, 128, 128); max-width: 100%; padding-left: 25px; line-height: 1; padding-bottom: 0px;' id='textSignature'><span id='fullName1' style='font-family:Biennale;color:#0073e6;font-weight:bold;font-size:13px'>Iosif-Alexandru Vrencean</span><br><span id='jobtitle1' style='font-family:Biennale;color:#000;font-size:11px'>Global Reporting Analyst</span><br><span style='font-family:Biennale;color:#000;font-size:11px' id='address1'>Atos IT Solutions and Services EOOD – 2 Knyaginya Maria Louis\a Blvd Sofia </span><span style='font-family:Biennale;color:#000;font-size:11px' id='addressline2span'> – 1000 – Bulgaria</span><span style='font-family:Biennale;color:#000;font-size:11px' id='country1'> – Bulgaria</span><br><a href='https://atos.net/bg' target='_blank' style='font-family:Biennale;font-size:11px;text-decoration:underline;color:#0073e6'>atos.net/bg</a><br><span id='idWrapIcons' class='wrapIcons' style='font-size:22px;margin-top:6px'><span><a target='_blank' href='https://www.linkedin.com/in/iosif-alexandru-vrencean-3747211b8/'><img src='https://atos.net/content/email-signature/linkedin.png' width='15' height='15' alt='LinkedIn icon'></a></span>&nbsp;<span><a title='LinkedIn' href='https://twitter.com/atos' target='_blank'><img src='https://atos.net/content/email-signature/twitter.png' width='15' height='15' alt='Twitter icon'></a></span>&nbsp;<span><a href='https://www.facebook.com/Atos/' target='_blank'><img src='https://atos.net/content/email-signature/facebook.png' width='15' height='15' alt='Facebook icon'></a></span>&nbsp;<span><a href='https://www.instagram.com/atosinside/' target='_blank'><img src='https://atos.net/content/email-signature/instagram.png' width='15' height='15' alt='Instagram icon'></a></span>&nbsp;<span><a href='https://www.youtube.com/user/Atos' target='_blank'><img src='https://atos.net/content/email-signature/youtube.png' width='15' height='15' alt='Youtube icon'></a></span>&nbsp;<span><a href='https://www.xing.com/companies/atosorigin' target='_blank'><img src='https://atos.net/content/email-signature/xing.png' width='15' height='15' alt='Xing icon'></a></span>&nbsp;</span></td></tr></tbody></table>"
    }

    "Report3" = @{
        "To" = "iosifa286@gmail.com"
        "CC" = ""
        "BCC" = ""
        "Subject" = "Report3"
        "Body" = "<div style='font-size:9pt; font-family:Biennale'>Hi Ohm,<br><br>Please find attached this week batch 1 of GOA offers for APAC, Growth Markets!<br>Also please check/refer On Hold Sheet created separately.<br><br></div><br><table cellpadding='0' cellspacing='0' border='0' width='100%' id='tablePreview' style='display: table; margin-bottom: 0px; transition: 300ms; font-size:9pt; font-family:Biennale;'><tbody><tr><td style='margin:0 0 15px 0;padding:0' id='politeSignature' colspan='2'><div id='polite1' style='margin:10px 0;font-family:Biennale;color:#000;font-size:9pt'>Kind regards,</div></td></tr><tr><td id='logoSignature' style='vertical-align: middle; padding-right: 20px; width: 150px;'><span><img width='150' src='https://atos.net/content/email-signature/euro-2024-atos.png' id='logo' alt='equensWorldline logo'></span></td><td width='100%' style='border-left: 2px solid rgb(128, 128, 128); max-width: 100%; padding-left: 25px; line-height: 1; padding-bottom: 0px;' id='textSignature'><span id='fullName1' style='font-family:Biennale;color:#0073e6;font-weight:bold;font-size:13px'>Iosif-Alexandru Vrencean</span><br><span id='jobtitle1' style='font-family:Biennale;color:#000;font-size:11px'>Global Reporting Analyst</span><br><span style='font-family:Biennale;color:#000;font-size:11px' id='address1'>Atos IT Solutions and Services EOOD – 2 Knyaginya Maria Louis\a Blvd Sofia </span><span style='font-family:Biennale;color:#000;font-size:11px' id='addressline2span'> – 1000 – Bulgaria</span><span style='font-family:Biennale;color:#000;font-size:11px' id='country1'> – Bulgaria</span><br><a href='https://atos.net/bg' target='_blank' style='font-family:Biennale;font-size:11px;text-decoration:underline;color:#0073e6'>atos.net/bg</a><br><span id='idWrapIcons' class='wrapIcons' style='font-size:22px;margin-top:6px'><span><a target='_blank' href='https://www.linkedin.com/in/iosif-alexandru-vrencean-3747211b8/'><img src='https://atos.net/content/email-signature/linkedin.png' width='15' height='15' alt='LinkedIn icon'></a></span>&nbsp;<span><a title='LinkedIn' href='https://twitter.com/atos' target='_blank'><img src='https://atos.net/content/email-signature/twitter.png' width='15' height='15' alt='Twitter icon'></a></span>&nbsp;<span><a href='https://www.facebook.com/Atos/' target='_blank'><img src='https://atos.net/content/email-signature/facebook.png' width='15' height='15' alt='Facebook icon'></a></span>&nbsp;<span><a href='https://www.instagram.com/atosinside/' target='_blank'><img src='https://atos.net/content/email-signature/instagram.png' width='15' height='15' alt='Instagram icon'></a></span>&nbsp;<span><a href='https://www.youtube.com/user/Atos' target='_blank'><img src='https://atos.net/content/email-signature/youtube.png' width='15' height='15' alt='Youtube icon'></a></span>&nbsp;<span><a href='https://www.xing.com/companies/atosorigin' target='_blank'><img src='https://atos.net/content/email-signature/xing.png' width='15' height='15' alt='Xing icon'></a></span>&nbsp;</span></td></tr></tbody></table>"
    }
}

$comboBox.Add_SelectedIndexChanged({
    $selectedTemplate = $comboBox.SelectedItem
    if ($templates.ContainsKey($selectedTemplate)) {
        $form.Controls["To"].Text = $templates[$selectedTemplate]["To"]
        $form.Controls["CC"].Text = $templates[$selectedTemplate]["CC"]
        $form.Controls["BCC"].Text = $templates[$selectedTemplate]["BCC"]
        $form.Controls["Subject"].Text = $templates[$selectedTemplate]["Subject"]
        $form.Controls["Body"].Text = $templates[$selectedTemplate]["Body"]
        $form.Controls["Attachment"].Text = $templates[$selectedTemplate]["Attachment"]
    }
})

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(140, 270)
$listBox.Size = New-Object System.Drawing.Size(400, 60)
$form.Controls.Add($listBox)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "Remove Attachment"
$removeButton.Location = New-Object System.Drawing.Point(550, 270)
$removeButton.Size = New-Object System.Drawing.Size(120, 30)
$removeButton.Add_Click({
    if ($listBox.SelectedItem) {
        Write-Host "Selected item: $($listBox.SelectedItem)"
        $additionalAttachments.Remove($listBox.SelectedItem) | Out-Null
        $listBox.Items.Remove($listBox.SelectedItem)
        Write-Host "Item removed."
    } else {
        Write-Host "No item selected."
    }
})
$form.Controls.Add($removeButton)

$addAttachmentButton = New-Object System.Windows.Forms.Button
$addAttachmentButton.Text = "Add Attachment"
$addAttachmentButton.Location = New-Object System.Drawing.Point(140, 230)
$addAttachmentButton.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($addAttachmentButton)

$additionalAttachments = New-Object System.Collections.ArrayList

$addAttachmentButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $openFileDialog.FileNames) {
            $additionalAttachments.Add($file) | Out-Null
            $listBox.Items.Add($file)
        }
    }
})

$sendButton = New-Object System.Windows.Forms.Button
$sendButton.Text = "Send Email"
$sendButton.Location = New-Object System.Drawing.Point(250, 400)
$sendButton.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($sendButton)

$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Text = "Preview Email"
$previewButton.Location = New-Object System.Drawing.Point(360, 400)
$previewButton.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($previewButton)

$sendButton.Add_Click({
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $form.Controls["To"].Text
    $mail.CC = $form.Controls["CC"].Text
    $mail.BCC = $form.Controls["BCC"].Text
    $mail.Subject = $form.Controls["Subject"].Text
    $mail.HTMLBody = $form.Controls["Body"].Text
    if ($form.Controls["Attachment"].Text -ne "") {
        $mail.Attachments.Add($form.Controls["Attachment"].Text)
    }
    foreach ($attachment in $additionalAttachments) {
        $mail.Attachments.Add($attachment)
    }
    $mail.Send()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    Remove-Variable outlook
    [System.Windows.Forms.MessageBox]::Show("Email sent successfully!")

    $additionalAttachments.Clear()
    $listBox.Items.Clear()
})

$previewButton.Add_Click({
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = $form.Controls["To"].Text
    $mail.CC = $form.Controls["CC"].Text
    $mail.BCC = $form.Controls["BCC"].Text
    $mail.Subject = $form.Controls["Subject"].Text
    $mail.HTMLBody = $form.Controls["Body"].Text
    if ($form.Controls["Attachment"].Text -ne "") {
        $mail.Attachments.Add($form.Controls["Attachment"].Text)
    }
    foreach ($attachment in $additionalAttachments) {
        $mail.Attachments.Add($attachment)
    }
    $mail.Display()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    Remove-Variable outlook
})

$excelButton = New-Object System.Windows.Forms.Button
$excelButton.Text = "Upload to refresh pivots"
$excelButton.Size = New-Object System.Drawing.Size(150, 50)
$excelButton.Location = New-Object System.Drawing.Point(470, 400)
$form.Controls.Add($excelButton)

$excelButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "Excel Files|*.xlsx"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $excelFilePath = $fileDialog.FileName

        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Open($excelFilePath)

        $refreshCount = 0
        $refreshLimit = 2
        $refreshInterval = 5

        while ($refreshCount -lt $refreshLimit) {
            Start-Sleep -Seconds $refreshInterval
            foreach ($sheet in $workbook.Sheets) {
                foreach ($pivotTable in $sheet.PivotTables()) {
                    $pivotTable.RefreshTable()
                }
            }
            $refreshCount++
        }

        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

        [System.Windows.Forms.MessageBox]::Show("Pivot tables refreshed successfully!")
    }
})

$uploadFileButton = New-Object System.Windows.Forms.Button
$uploadFileButton.Text = "Upload File"
$uploadFileButton.Size = New-Object System.Drawing.Size(150, 50)
$uploadFileButton.Location = New-Object System.Drawing.Point(250, 230)
$form.Controls.Add($uploadFileButton)

$uploadFileButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "All Files|*.*"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFilePath = $fileDialog.FileName

        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.SelectedPath = "C:\Users\a880862\OneDrive - ATOS\Data Integrity and Reports"
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $destinationFolderPath = $folderDialog.SelectedPath

            $destinationFilePath = Join-Path -Path $destinationFolderPath -ChildPath (Split-Path -Leaf $selectedFilePath)
            Copy-Item -Path $selectedFilePath -Destination $destinationFilePath

            [System.Windows.Forms.MessageBox]::Show("File uploaded successfully to $destinationFolderPath")
        }
    }
})

$form.ShowDialog()