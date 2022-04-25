# ----------------------------------------------------------
# Copyright © Novembre 2019 - Éric Gorin
# ----------------------------------------------------------
# Chargement des librairies de code VB, Forms et Drawing
# ----------------------------------------------------------
# Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms    
Add-Type -AssemblyName System.Drawing

# ----------------------------------------------------------
# Chargement des utilisateurs locaux dans la combo box
# ----------------------------------------------------------
function ChargeUtilisateurs($param1)   
{  
#-------------------------------------------------------------------    
#    if($param1 -like "*param*") {
#       [System.Windows.Forms.MessageBox]::Show($param1, "test")
#    }else {
#        [System.Windows.Forms.MessageBox]::Show("Bad param", "test")
#    }
#-------------------------------------------------------------------    
   $ComboBox.ResetText()
   $ComboBox.Items.Clear()
   $Utilisateurs = Get-LocalUser | ? {$_.Enabled -eq "True"}
   Foreach ($Utilisateur in $Utilisateurs) 
   {
    $ComboBox.Items.Add($Utilisateur);
   } 
}

# ----------------------------------------------------------
# Chargement des groupes locaux dans la combo box
# ----------------------------------------------------------
function ChargeGroupes   
{  
   #[System.Windows.Forms.MessageBox]::Show("ChargeUtilisateurs" , "test") 
   $ComboBoxGroupes.ResetText()
   $ComboBoxGroupes.Items.Clear()
   $Groupes = Get-LocalGroup | Where-Object {$_.Name -like '*utili*' -or $_.Description -like '*membres*'}
   Foreach ($Groupe in $Groupes) 
   {
    $ComboBoxGroupes.Items.Add($Groupe);
   } 
}

# ------------------------------------------
# Création d'un obget pour centrer une form
# ------------------------------------------
$CentreEcran = [System.Windows.Forms.FormStartPosition]::CenterScreen;

# ----------------------------------
# Création d'une police (font)
# ----------------------------------
$Police = New-Object System.Drawing.Font("Times New Roman",20,[System.Drawing.FontStyle]::Italic)
$fSizeBold =  New-Object System.Drawing.Font("GenericSansSerif", 10,[System.Drawing.FontStyle]::Bold)

# ----------------------------------
# Création de la form
# ----------------------------------
$form_principal = New-Object System.Windows.Forms.Form
$form_principal.Text ='Mon interface écran'
$form_principal.Width = 600
$form_principal.Height = 400
$form_principal.AutoSize = $true
$form_principal.StartPosition = $CentreEcran


# ----------------------------------
# Création d'un label titre
#-----------------------------------
$Mon_label_titre = New-Object System.Windows.Forms.Label
$Mon_label_titre.Text = "Gestion des utilisateurs"
$Mon_label_titre.Location  = New-Object System.Drawing.Point(130,10)
$Mon_label_titre.AutoSize = $true
$Mon_label_titre.Font = $Police
#-----------------------------------
# Ajout de mon label dans ma form 
#-----------------------------------
$form_principal.Controls.Add($Mon_label_titre)

# ----------------------------------------------
# Création d'un label de liste des utilisateurs
#----------------------------------------------
$Mon_label_liste = New-Object System.Windows.Forms.Label
$Mon_label_liste.Text = "Liste des utilisateurs"
$Mon_label_liste.Location  = New-Object System.Drawing.Point(325,60)
$Mon_label_liste.AutoSize = $true
$Mon_label_liste.Font = $fSizeBold
#-----------------------------------
# Ajout de mon label 2 dans ma form 
#-----------------------------------
$form_principal.Controls.Add($Mon_label_liste)

#-----------------------------------
# Création d'une compo box 
#-----------------------------------
$ComboBox = New-Object System.Windows.Forms.ComboBox
$ComboBox.Width = 300
#$ComboBox.DropDownStyle=[System.Windows.Forms.ComboBoxStyle]::DropDownList
$ComboBox.DropDownStyle='DropDownList'
ChargeUtilisateurs -param1 "Eric"
#$Utilisateurs = Get-LocalUser | ? {$_.Enabled -eq "True"}
#Foreach ($Utilisateur in $Utilisateurs) 
#{
#$ComboBox.Items.Add($Utilisateur);
#}
$ComboBox.Location  = New-Object System.Drawing.Point(250,90)
#-----------------------------------
# Ajout de ma compo box dans ma form  
#-----------------------------------
$form_principal.Controls.Add($ComboBox)


# ----------------------------------
# Création d'un label Ajout utilisateur
#-----------------------------------
$Mon_label_ajout = New-Object System.Windows.Forms.Label
$Mon_label_ajout.Text = "Ajout d'un utilisateur"
$Mon_label_ajout.Location  = New-Object System.Drawing.Point(20,60)
$Mon_label_ajout.AutoSize = $true
$Mon_label_ajout.Font = $fSizeBold
#-----------------------------------
# Ajout de mon label Ajout utilisateur dans ma form 
#-----------------------------------
$form_principal.Controls.Add($Mon_label_ajout)


#-----------------------------------
# Création d'un bouton d'ajout 
#-----------------------------------
$BTAjout = New-Object System.Windows.Forms.Button
$BTAjout.Location  = New-Object System.Drawing.Point(65,120)
$BTAjout.Text = "Ajout"
#-----------------------------------
# Ajout de mon bouton dans ma form  
#-----------------------------------
$form_principal.Controls.Add($BTAjout)


#----------------------------------------------------
# Création d'une text box pour le nouvel utilisateur
#----------------------------------------------------
$Ma_TextBox = New-Object System.Windows.Forms.TextBox
$Ma_TextBox.Location  = New-Object System.Drawing.Point(55,90)

$Ma_TextBox_KeyDown = [System.Windows.Forms.KeyEventHandler]{
	 [System.Windows.Forms.MessageBox]::Show("Key_code", "test")
     if ($_.KeyCode -eq 'Enter')
	{
		$BTAjout.Invoke()
		$BTAjout.PerformClick()
	}
}

#-----------------------------------
# Ajout de ma text box dans ma form  
#-----------------------------------
$form_principal.Controls.Add($Ma_TextBox)

$Ma_TextBox_KeyDown = {}
$Ma_TextBox_KeyDown = [System.Windows.Forms.KeyEventHandler]{}
$Ma_TextBox_KeyDown = [System.Windows.Forms.KeyEventHandler]{
    if ($_.KeyCode ==  Keys.Enter)
    {
        [System.Windows.Forms.MessageBox]::Show("Key_code", "test")
        $BTAjout.PerformClick()
    }
}


#--------------------------------------------------
# Création d'un bouton de supression d'utilisateur
#--------------------------------------------------
$BTSupprime = New-Object System.Windows.Forms.Button
$BTSupprime.Location  = New-Object System.Drawing.Point(310,120)
$BTSupprime.Text = "Supprimer un utilisateur"
$BTSupprime.AutoSize = $true
#-----------------------------------
# Ajout de mon bouton dans ma form  
#-----------------------------------
$form_principal.Controls.Add($BTSupprime)

#-------------------------------------------------------------------
# Création d'un evenement pour le bouton de supression d'utilisateur
#-------------------------------------------------------------------
$BTSupprime.Add_Click( 
 {
  #[System.Windows.Forms.MessageBox]::Show($ComboBox.SelectedItem.ToString() , "supression test")    
  Remove-LocalUser -Name $ComboBox.SelectedItem.ToString()
  ChargeUtilisateurs
 }
)


#--------------------------------------------------
# Création d'un bouton de modification de mot de passe
#--------------------------------------------------
$BTModif = New-Object System.Windows.Forms.Button
$BTModif.Location  = New-Object System.Drawing.Point(305,170)
$BTModif.Text = "Modifier un mot de passe"
$BTModif.AutoSize = $true
#-----------------------------------
# Ajout de mon bouton dans ma form  
#-----------------------------------
$form_principal.Controls.Add($BTModif)


#--------------------------------------------------
# Création d'une compo box qui affiche les groupes
#-------------------------------------------------
$ComboBoxGroupes = New-Object System.Windows.Forms.ComboBox
$ComboBoxGroupes.Width = 300
ChargeGroupes
$ComboBoxGroupes.Location  = New-Object System.Drawing.Point(250,220)
$ComboBoxGroupes.DropDownStyle = 'DropDownList'
#-----------------------------------
# Ajout de ma compo box dans ma form  
#-----------------------------------
$form_principal.Controls.Add($ComboBoxGroupes)


#-----------------------------------------------------------------------------
# Création d'un bouton d'association d'un groupe a un utilisateur sélectionnée
#-----------------------------------------------------------------------------
$BTAssoc = New-Object System.Windows.Forms.Button
$BTAssoc.Location  = New-Object System.Drawing.Point(325,260)
$BTAssoc.Text = "Associé un groupe"
$BTAssoc.AutoSize = $true
#-----------------------------------
# Ajout de mon bouton dans ma form  
#-----------------------------------
$form_principal.Controls.Add($BTAssoc)


#-------------------------------------------------------------------
# Création d'un evenement pour le bouton d'association d'un groupe
#-------------------------------------------------------------------
$BTAssoc.Add_Click( 
 {
  #[System.Windows.Forms.MessageBox]::Show($ComboBox.SelectedItem.ToString() + " " + $ComboBoxGroupes.SelectedItem.ToString(), "association test")    
  Add-LocalGroupMember -Group $ComboBoxGroupes.SelectedItem.ToString() -Member $ComboBox.SelectedItem.ToString() 
  $infoGroup = Get-LocalGroupMember -Group $ComboBoxGroupes.SelectedItem.ToString() | Select Name | where Name -like ('*' + $ComboBox.SelectedItem.ToString() + '*') | Format-Table -HideTableHeaders 
  $infoGroup = $infoGroup | Out-String
  [System.Windows.Forms.MessageBox]::Show($infoGroup + "est maintenant dans le groupe " + $ComboBoxGroupes.SelectedItem.ToString(), "Group utilisateur info")   
 }
)

#-----------------------------------------------------------------------
# Création d'un evenement pour le bouton de modification de mot de passe
#-----------------------------------------------------------------------
$BTModif.Add_Click( 
 {
  #[System.Windows.Forms.MessageBox]::Show($ComboBox.SelectedItem.ToString() , "supression test")          
   # -----------------------------------------
   # Création de la form pour le mot de passe
   # ---------------------------------------------
   $form_modif_pwd = New-Object System.Windows.Forms.Form
   $form_modif_pwd.Text ='modifier le mot de passe'
   $form_modif_pwd.Width = 350
   $form_modif_pwd.Height = 100
   $form_modif_pwd.AutoSize = $true  
   $form_modif_pwd.StartPosition = $CentreEcran
   #--------------------------------------------
   # Création d'une text box pour modifier le mot de passe
   #--------------------------------------------
   $modif_pwd = New-Object System.Windows.Forms.MaskedTextBox
   $modif_pwd.Width = 200
   $modif_pwd.PasswordChar = '?'
   $modif_pwd.Location  = New-Object System.Drawing.Point(10,10)    
   #---------------------------------------------------------------------
   # Ajout de ma text box (modif mot de passe) dans ma form $form_modif_pwd  
   #--------------------------------------------------------------------
   $form_modif_pwd.Controls.Add($modif_pwd)    
   #-------------------------------------------------------
   # Création d'un bouton ok pour la modif du mot de passe 
   #------------------------------------------------------
   $BTOk2 = New-Object System.Windows.Forms.Button
   $BTOk2.Location  = New-Object System.Drawing.Point(225,08)
   $BTOk2.Text = "Ok"
   #---------------------------------------------------
   # Ajout de mon bouton dans ma form $form_modif_pwd  
   #--------------------------------------------------
   $form_modif_pwd.Controls.Add($BTOk2)
   #------------------------------------------------
   # Création d'un evenement pour le bouton d'ajout 
   #------------------------------------------------
   $BTOk2.Add_Click( 
    {    
    # ---------------------------------------------------------
    # changement du mot de passe pour l'utilisateur selectionné
    # ---------------------------------------------------------  
    $Utilisateur = Get-LocalUser -Name $ComboBox.SelectedItem.ToString()
    $Utilisateur | Set-LocalUser -Password  (ConvertTo-SecureString -AsPlainText $modif_pwd.Text -Force)
    ChargeUtilisateurs
    $Ma_TextBox.Text = ""  
    $form_modif_pwd.Close()
    }
   )
   #---------------------------------------------------------------------
   # Affichage de ma form $form_mot_de_passe   
   #--------------------------------------------------------------------
   $form_modif_pwd.ShowDialog() 
         
 }
)


#------------------------------------------------
# Création d'un evenement pour le bouton d'ajout 
#------------------------------------------------
 $BTAjout.Add_Click( 
   {    
   # -----------------------------------------
   # Création de la form pour le mot de passe
   # ---------------------------------------------
   $form_mot_de_passe = New-Object System.Windows.Forms.Form
   $form_mot_de_passe.Text ='mot de passe'
   $form_mot_de_passe.Width = 350
   $form_mot_de_passe.Height = 100
   $form_mot_de_passe.AutoSize = $true  
   $form_mot_de_passe.StartPosition = $CentreEcran
   #--------------------------------------------
   # Création d'une text box pour mot de passe
   #--------------------------------------------
   $pwd = New-Object System.Windows.Forms.MaskedTextBox
   $pwd.Width = 200
   $pwd.PasswordChar = '*'
   $pwd.Location  = New-Object System.Drawing.Point(10,10) 
   #---------------------------------------------------------------------
   # Ajout de ma text box (mot de passe) dans ma form $form_mot_de_passe   
   #--------------------------------------------------------------------
   $form_mot_de_passe.Controls.Add($pwd)    
   #----------------------------------------------
   # Création d'un bouton ok pour le mot de passe 
   #----------------------------------------------
   $BTOk = New-Object System.Windows.Forms.Button
   $BTOk.Location  = New-Object System.Drawing.Point(225,08)
   $BTOk.Text = "Ok"
   #-----------------------------------
   # Ajout de mon bouton dans ma form  
   #-----------------------------------
   $form_mot_de_passe.Controls.Add($BTOk)
      
    #------------------------------------------------
    # Création d'un evenement pour le bouton d'ajout 
    #------------------------------------------------
    $BTOk.Add_Click( 
       {    
       # -------------------------------
       # Création du nouvel utilisateur 
       # -------------------------------
       #[System.Windows.Forms.MessageBox]::Show($pwd.Text , "Password test")
        New-LocalUser -Name  $Ma_TextBox.Text -Description "Cours logiciel" -Password  (ConvertTo-SecureString -AsPlainText $pwd.Text -Force)
        ChargeUtilisateurs
        $Ma_TextBox.Text = ""  
        $form_mot_de_passe.Close()
       }
    )
   #---------------------------------------------------------------------
   # Affichage de ma form $form_mot_de_passe   
   #--------------------------------------------------------------------
   $form_mot_de_passe.ShowDialog()    
   }
 )
 


#---------------------------------------------------------------------
# Affichage de ma form $form_mot_de_passe   
#--------------------------------------------------------------------
$form_principal.ShowDialog()




