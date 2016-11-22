Function Search-MailboxGUI
{
    #Calling the Assemblies
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  | Out-Null

    # Create the Conatainer Form  to place the Labels and Textboxes
    $Form = New-Object “System.Windows.Forms.Form”;
    $Form.Width = 500;
    $Form.Height = 400;
    $Form.Text = 'Enter the required details'
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Form.TopMost = $true

    #Define Label1
    $Label1 = New-Object “System.Windows.Forms.Label”;
    $Label1.Left = 10;
    $Label1.Top = 15;
    $Label1.Text = 'Recipient Email';
    #Define Label2
    $Label2 = New-Object “System.Windows.Forms.Label”;
    $Label2.Left = 10;
    $Label2.Top = 40;
    $Label2.Text = 'Sender Email';
    #Define Label3
    $Label3 = New-Object “System.Windows.Forms.Label”;
    $Label3.Left = 10;
    $Label3.Top = 65;
    $Label3.Width = 180
    $Label3.Text = 'Search Keyword';

    #Define Label4
    $Label4 = New-Object “System.Windows.Forms.Label”;
    $Label4.Left = 10;
    $Label4.Top = 115;
    $Label4.Width = 180
    $Label4.Text = 'Start Date [MM/DD/YYYY]';
    
    #Define Label5
    $Label5 = New-Object “System.Windows.Forms.Label”;
    $Label5.Left = 10;
    $Label5.Top = 140;
    $Label5.Width = 180
    $Label5.Text = 'End Date [MM/DD/YYYY]';
#Define Label6
    $Label6 = New-Object “System.Windows.Forms.Label”;
    $Label6.Left = 10;
    $Label6.Top = 165;
    $Label6.Width = 180
    $Label6.Text = 'Delete Email';


    #Define Label7
    $Label7 = New-Object “System.Windows.Forms.Label”;
    $Label7.Left = 10;
    $Label7.Top = 190;
    $Label7.Width = 180
    $Label7.Text = 'Admin Email to keep a copy';


    #Define Label8
    $Label8 = New-Object “System.Windows.Forms.Label”;
    $Label8.Left = 10;
    $Label8.Top = 90;
    $Label8.Width = 180
    $Label8.Text = 'Search Only in Subject line';

    #Define TextBox1 for input
    $TextBox1 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox1.Left = 200;
    $TextBox1.Top = 15;
    $TextBox1.width = 250;
#Define TextBox2 for input
    $TextBox2 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox2.Left = 200;
    $TextBox2.Top = 40;
    $TextBox2.width = 250;

    #Define TextBox3 for input
    $TextBox3 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox3.Left = 200;
    $TextBox3.Top = 65;
    $TextBox3.width = 250;

    #Define Textbox4 for input
    $TextBox4 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox4.Left = 200;
    $TextBox4.Top = 115;
    $TextBox4.width = 250;

    #Define Textbox5 for input
    $TextBox5 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox5.Left = 200;
    $TextBox5.Top = 140;
    $TextBox5.width = 250;

    #Define Radio Button
    $CheckBox = New-Object System.Windows.Forms.CheckBox
    $CheckBox.Left = 200
    $CheckBox.Top = 165
#Define Radio Button
    $CheckBox2 = New-Object System.Windows.Forms.CheckBox
    $CheckBox2.Left = 200
    $CheckBox2.Top = 90

    #Define Textbox6 for input
    $TextBox6 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox6.Left = 200;
    $TextBox6.Top = 190;
    $TextBox6.width = 250;

    #Define OK Button
    $OKbutton = New-Object “System.Windows.Forms.Button”;
    $OKbutton.Left = 10;
    $OKbutton.Top = 220;
    $OKbutton.Width = 100;
    $OKbutton.Text = “SEARCH”;


    ############# This is when you have to close the Form after getting values
    $eventHandler = [System.EventHandler]{
    $Form.Close();
    };

    $OKbutton.Add_Click($eventHandler) ;
#############Add controls to all the above objects defined
    $Form.Controls.Add($OKbutton);
    $Form.Controls.Add($Label1);
    $Form.Controls.Add($Label2);
    $Form.Controls.Add($Label3);
    $Form.Controls.Add($Label4);
    $Form.Controls.Add($Label5);
    $Form.Controls.Add($Label6);
    $Form.Controls.Add($Label7);
    $Form.Controls.Add($Label8);
    $Form.Controls.Add($CheckBox2);
    $Form.Controls.Add($TextBox1);
    $Form.Controls.Add($TextBox2);
    $Form.Controls.Add($TextBox3);
    $Form.Controls.Add($TextBox4);
    $Form.Controls.Add($TextBox5);    
    $Form.Controls.Add($CheckBox);
    $Form.Controls.Add($TextBox6);    
    $Form.ShowDialog()|Out-Null


    #Extracting User data into variables
    $TargetEmail = $TextBox1.Text; 
    $AdminEmail =$TextBox6.Text
    $SenderEmail=$TextBox2.Text;
    $Keyword=$TextBox3.Text; 
    $StartDate= $TextBox4.Text; 
    $EndDate=$TextBox5.Text

    #CheckBox2 = Seacrh Keyword in Subject only
    If($CheckBox2.Checked)
    {
        $Subject='subject:'
    }
    else
    {
        $Subject=$null
    }

    #Iterate all recipient Mailboxes and search the Keyword in it
    Foreach($Target in $TargetEmail.Split(';'))
    {
        #Check the Condition if Delete email checkbox is checked or not
        If($CheckBox.Checked)
        {
            Search-Mailbox -Identity $Target -TargetMailbox $AdminEmail -TargetFolder EmailCaptures -LogLevel full -SearchQuery "$Subject$Keyword AND from:$SenderEmail and received:>$([datetime]$StartDate) and received:<$([datetime]$EndDate)" -DeleteContent -Confirm:$false
        }
        Else
        {
            Search-Mailbox -Identity $Target -TargetMailbox $AdminEmail -TargetFolder EmailCaptures -LogLevel full -SearchQuery "$Subject$Keyword AND from:$SenderEmail and received:>$([datetime]$StartDate) and received:<$([datetime]$EndDate)"
        }

    }
    }

Function Create-Form
{
    #Calling the Assemblies
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  | Out-Null

    # Create the Conatainer Form  to place the Labels and Textboxes
    $Form = New-Object “System.Windows.Forms.Form”;
    $Form.Width = 500;
    $Form.Height = 400;
    $Form.Text = 'Enter the required details'
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $Form.TopMost = $true

    #Define Label1
    $Label1 = New-Object “System.Windows.Forms.Label”;
    $Label1.Left = 10;
    $Label1.Top = 15;
    $Label1.Text = 'Recipient Email';
    #Define Label2
    $Label2 = New-Object “System.Windows.Forms.Label”;
    $Label2.Left = 10;
    $Label2.Top = 40;
    $Label2.Text = 'Sender Email';
    #Define Label3
    $Label3 = New-Object “System.Windows.Forms.Label”;
    $Label3.Left = 10;
    $Label3.Top = 65;
    $Label3.Width = 180
    $Label3.Text = 'Search Keyword';

    #Define Label4
    $Label4 = New-Object “System.Windows.Forms.Label”;
    $Label4.Left = 10;
    $Label4.Top = 115;
    $Label4.Width = 180
    $Label4.Text = 'Start Date [MM/DD/YYYY]';
    
    #Define Label5
    $Label5 = New-Object “System.Windows.Forms.Label”;
    $Label5.Left = 10;
    $Label5.Top = 140;
    $Label5.Width = 180
    $Label5.Text = 'End Date [MM/DD/YYYY]';
#Define Label6
    $Label6 = New-Object “System.Windows.Forms.Label”;
    $Label6.Left = 10;
    $Label6.Top = 165;
    $Label6.Width = 180
    $Label6.Text = 'Delete Email';


    #Define Label7
    $Label7 = New-Object “System.Windows.Forms.Label”;
    $Label7.Left = 10;
    $Label7.Top = 190;
    $Label7.Width = 180
    $Label7.Text = 'Admin Email to keep a copy';


    #Define Label8
    $Label8 = New-Object “System.Windows.Forms.Label”;
    $Label8.Left = 10;
    $Label8.Top = 90;
    $Label8.Width = 180
    $Label8.Text = 'Search Only in Subject line';

    #Define TextBox1 for input
    $TextBox1 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox1.Left = 200;
    $TextBox1.Top = 15;
    $TextBox1.width = 250;
#Define TextBox2 for input
    $TextBox2 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox2.Left = 200;
    $TextBox2.Top = 40;
    $TextBox2.width = 250;

    #Define TextBox3 for input
    $TextBox3 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox3.Left = 200;
    $TextBox3.Top = 65;
    $TextBox3.width = 250;

    #Define Textbox4 for input
    $TextBox4 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox4.Left = 200;
    $TextBox4.Top = 115;
    $TextBox4.width = 250;

    #Define Textbox5 for input
    $TextBox5 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox5.Left = 200;
    $TextBox5.Top = 140;
    $TextBox5.width = 250;

    #Define Radio Button
    $CheckBox = New-Object System.Windows.Forms.CheckBox
    $CheckBox.Left = 200
    $CheckBox.Top = 165
#Define Radio Button
    $CheckBox2 = New-Object System.Windows.Forms.CheckBox
    $CheckBox2.Left = 200
    $CheckBox2.Top = 90

    #Define Textbox6 for input
    $TextBox6 = New-Object “System.Windows.Forms.TextBox”;
    $TextBox6.Left = 200;
    $TextBox6.Top = 190;
    $TextBox6.width = 250;

    #Define OK Button
    $OKbutton = New-Object “System.Windows.Forms.Button”;
    $OKbutton.Left = 10;
    $OKbutton.Top = 220;
    $OKbutton.Width = 100;
    $OKbutton.Text = “SEARCH”;


    ############# This is when you have to close the Form after getting values
    $eventHandler = [System.EventHandler]{
    $Form.Close();
    };

    $OKbutton.Add_Click($eventHandler) ;
#############Add controls to all the above objects defined
    $Form.Controls.Add($OKbutton);
    $Form.Controls.Add($Label1);
    $Form.Controls.Add($Label2);
    $Form.Controls.Add($Label3);
    $Form.Controls.Add($Label4);
    $Form.Controls.Add($Label5);
    $Form.Controls.Add($Label6);
    $Form.Controls.Add($Label7);
    $Form.Controls.Add($Label8);
    $Form.Controls.Add($CheckBox2);
    $Form.Controls.Add($TextBox1);
    $Form.Controls.Add($TextBox2);
    $Form.Controls.Add($TextBox3);
    $Form.Controls.Add($TextBox4);
    $Form.Controls.Add($TextBox5);    
    $Form.Controls.Add($CheckBox);
    $Form.Controls.Add($TextBox6);    
    $Form.ShowDialog()|Out-Null


    #Extracting User data into variables
    $TargetEmail = $TextBox1.Text; 
    $AdminEmail =$TextBox6.Text
    $SenderEmail=$TextBox2.Text;
    $Keyword=$TextBox3.Text; 
    $StartDate= $TextBox4.Text; 
    $EndDate=$TextBox5.Text

    #CheckBox2 = Seacrh Keyword in Subject only
    If($CheckBox2.Checked)
    {
        $Subject='subject:'
    }
    else
    {
        $Subject=$null
    }

    #Iterate all recipient Mailboxes and search the Keyword in it
    Foreach($Target in $TargetEmail.Split(';'))
    {
        #Check the Condition if Delete email checkbox is checked or not
        If($CheckBox.Checked)
        {
            Search-Mailbox -Identity $Target -TargetMailbox $AdminEmail -TargetFolder EmailCaptures -LogLevel full -SearchQuery "$Subject$Keyword AND from:$SenderEmail and received:>$([datetime]$StartDate) and received:<$([datetime]$EndDate)" -DeleteContent -Confirm:$false
        }
        Else
        {
            Search-Mailbox -Identity $Target -TargetMailbox $AdminEmail -TargetFolder EmailCaptures -LogLevel full -SearchQuery "$Subject$Keyword AND from:$SenderEmail and received:>$([datetime]$StartDate) and received:<$([datetime]$EndDate)"
        }

    }
    }

Search-MailboxGUI
