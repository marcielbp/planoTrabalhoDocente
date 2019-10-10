// Google Doc id from the document template
// (Get ids from the URL)
var SOURCE_TEMPLATE = "1yzdDPoKKD4-4x6PDxI1dD0QV3_yb18HU6w0k8-czRxY"; // A variável recebe como valor o ID do template

// In which spreadsheet we have all the customer data
var TARGET_SHEET = "1TI43nlyrU03LtuXfFPEGjL3MC-Z0EaWGfUJFFcHUGHo"; // A variável recebe como valor o ID da planilha

// In which Google Drive we toss the target documents
var TARGET_FOLDER = "1waXe6Cg9izU6XfQUxvpNWcXVxgkAX_CX";

function myFunction() {

  var source = DriveApp.getFileById(SOURCE_TEMPLATE); //Apenas lê o arquivo, sem liberar métodos específicos para o Docs
  var sourceSheet = DriveApp.getFileById(TARGET_SHEET); // Apenas lê o arquivo, sem liberar métodos específicos para o Spreadsheets
  var sheet = SpreadsheetApp.openById(sourceSheet.getId()); // Abre a planilha de acordo com o ID recebido
  var eachFile, idToDLET, myFolder, rtrnFromDLET, thisFile;

  myFolder = DriveApp.getFolderById(TARGET_FOLDER);
  thisFile = myFolder.getFiles();
  while (thisFile.hasNext()) 
  {//If there is another element in the iterator
    eachFile = thisFile.next();
    idToDLET = eachFile.getId();
    Logger.log('idToDLET: ' + idToDLET);
    
    //rtrnFromDLET = Drive.Files.remove(idToDLET);
  }
  var data = sheet.getDataRange().getValues(); // A variável recebe os dados presentes na planilha
  for (var j = 1; j<data.length; j++)
  {
    if(data[j][27] == "OK")//Coluna AB da planilha
      {
        var newFile = source.makeCopy("S2019-2_PlTr_"+data[j][2]); //Faz uma cópia do template com o nome Resultado
        var targetFolder = DriveApp.getFolderById(TARGET_FOLDER);
        targetFolder.addFile(newFile);
        //newFile.addToFolder(targetFolder);
        var doc = DocumentApp.openById(newFile.getId()); //Abre o documento de acordo com o ID recebido
        //var now = new Date();
        var now = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy");
        var body = doc.getBody(); // A variável recebe o corpo do documento
        var ps = body.getParagraphs(); // A variável recebe os paragráfos do documento
    
        for(var i=0; i<ps.length; i++) 
        {
          var p = ps[i];
          var text = p.getText();
          p.replaceText("#NAME#", data[j][2]);
          p.replaceText("#SIAPE#", data[j][3]);
          p.replaceText("#REGTR#", data[j][4]);
          p.replaceText("#CLASS#", data[j][5]);
          
          p.replaceText("#HOR01#", data[j][6]);
          p.replaceText("#AT01#", "■ " + data[j][7]);
          p.replaceText("#CHAT01#", data[j][8]);
          
          p.replaceText("#HOR02#", data[j][9]);
          p.replaceText("#AT02#", "■ " + data[j][10]);
          p.replaceText("#CHAT02#", data[j][11]);
         
          p.replaceText("#HOR03#", data[j][12]);
          p.replaceText("#AT03#", "■ " + data[j][13]);
          p.replaceText("#CHAT03#", data[j][14]);

          p.replaceText("#HOR04#", data[j][15]);
          p.replaceText("#AT04#", "■ " + data[j][16]);
          p.replaceText("#CHAT04#", data[j][17]);

          p.replaceText("#HOR05#", data[j][18]);
          p.replaceText("#AT05#", "■ " + data[j][19]);
          p.replaceText("#CHAT05#", data[j][20]);
          
          p.replaceText("#HOR06#", data[j][21]);
          p.replaceText("#AT06#", "■ " + data[j][22]);
          p.replaceText("#CHAT06#", data[j][23]);
          
          p.replaceText("#DATAAPROVACAO", data[j][24]);
          
          p.replaceText(";, ", "\n■ ");
          p.replaceText(";", "\n");
        }
        
      }
  }
  // A estrutura de repetição acima é responsável por fazer a mudança de texto dentro do documento de acordo com a planilha
  // Ambos tendo sido especificados anteriormente
}
