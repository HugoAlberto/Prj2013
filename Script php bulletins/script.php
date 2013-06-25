<?
/*--------------------------------------------------------------------------------*\
**                           Envoi de bulletin par mail.                          **
**                             Création : Boulouk Hugo                            **
**                                 Le 2013/06/24                                  **
**                                Modification : --                               **
\*--------------------------------------------------------------------------------*/
// Aide à l'envoi de mail
require_once "phpmailer/class.phpmailer.php";
// Aide à l'exploration de fichier Excel
require "phpexel/Classes/PHPExcel.php";

/*--------------------------------------------------------------------------------*\
**                                   Variables                                    **
\*--------------------------------------------------------------------------------*/
// Mail émetteur
$mailFrom = 'info.stjo@orange.fr';
$mailFromName = 'Informatique - St Joseph Gap';
// Mail réponse (par defaut le même que l'émetreur)
$mailReply = $mailFrom;
// Objet du mail
$subject = utf8_decode('Bulletin scolaire de votre enfant (collège)');
// Nom du répertoire ou ce trouve les bulletins -- <4C/>
$dirname = '4C/';
// Nom du fichier .xls à explorer -- <nomDuFichier.xlsx>
$xlsxname = 'Liste_4c.xlsx';

/*--------------------------------------------------------------------------------*\
**                             Exploration fichier Excel                          **
\*--------------------------------------------------------------------------------*/
// Création de l'objet Reader pour un fichier Excel
$objReader = new PHPExcel_Reader_Excel2007();
// Permet de ne récupérer que les valeurs des cellules sans les propriétés de style 
$objReader->setReadDataOnly(true);
// Lecture du fichier
$objPHPExcel = $objReader->load($xlsxname);
// Permet de récupérer toutes les données
$rowIterator = $objPHPExcel->getActiveSheet()->getRowIterator();
$dir = opendir($dirname); 
while($file = readdir($dir)) 
{
  if($file != '.' && $file != '..' && !is_dir($dirname.$file))
  {
    $parentName = basename($file,".pdf");
    list($namePdf,$parentPdf) = explode("_", $parentName);
    $namePdf=strtoupper($namePdf);
    if($parentPdf == "papa")
    {
      $parentPdf = "Monsieur";
    }
    elseif($parentPdf == "maman")
    {
      $parentPdf = "Madame";
    }
    foreach($rowIterator as $row)
    {
      $cellIterator = $row->getCellIterator();
      $cellIterator->setIterateOnlyExistingCells(false);
      $rowIndex = $row->getRowIndex();
      $array_data[$rowIndex] = array('A'=>'', 'B'=>'','C'=>'','D'=>'');
      foreach ($cellIterator as $cell) 
      {
        if('A' == $cell->getColumn())
        {
          $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        } 
        else if('B' == $cell->getColumn())
        {
          $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        } 
        else if('C' == $cell->getColumn()) 
        {
          $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        } 
        else if('D' == $cell->getColumn()) 
        {
          $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
      }
      // affiche le contenu de la ligne pour la colonne x
      $FnameXl = $array_data[$rowIndex]['A'];
      $nameXl = $array_data[$rowIndex]['B'];
      $mailXl = $array_data[$rowIndex]['C'];
      $parentXl = $array_data[$rowIndex]['D'];

  /*--------------------------------------------------------------------------------*\
  **                              Nom du fichier pdf                                **
  \*--------------------------------------------------------------------------------*/
      if($namePdf == $FnameXl && $parentPdf == $parentXl)
      {
        //echo $FnameXl,$nameXl,$mailXl,$parentXl,"\n";
	$mailTo = $mailXl;
        if($parentPdf == "Monsieur")
        {
           $parentPdf = "papa";
        }
        elseif($parentPdf == "Madame")
        {
          $parentPdf = "maman";
        }
	$namePdf = strtolower($namePdf);
	$pieceJointe = $dirname.$namePdf."_".$parentPdf.".pdf";
	
/*--------------------------------------------------------------------------------*\
**                                  Envoi de mail                                 **
\*--------------------------------------------------------------------------------*/
	$mail = new PHPmailer();
	$mail->IsHTML(true);
	// Emetteur
	$mail->From = $mailFrom; 
	$mail->FromName = $mailFromName;
	// Destinataire
	$mail->AddAddress($mailTo);
	// e-Mail de réponse
	$mail->AddReplyTo($mailReply);
	// Sujet du mail
	$mail->Subject = $subject;
	// Corps du mail
	$mail->Body = '<html><body><center><font size=3>Le fichier est attaché ci-dessus</font><br></body></html>';
	// Ajout du fichier pdf 
	$mail->AddAttachment($pieceJointe); 
	// Si l'e-mail n'a pas été envoyé
	if(!$mail->Send())
	{ 
	  echo $mail->ErrorInfo;  
	} 
	// Sinon
	else
	{      
	  echo 'Mail envoyé avec succès',"\n"; 
	}
	// Fermeture de la connexion
	$mail->SmtpClose(); 
	unset($mail);
      }
    }
  }
}
closedir($dir);

?>
