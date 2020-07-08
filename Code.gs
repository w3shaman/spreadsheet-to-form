function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Buat Google Form', functionName: 'createQuiz'}
  ];
  spreadsheet.addMenu('Kuis', menuItems);
}

function createQuiz() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var nama_form = data[0][1];
  var deskripsi_form = data[1][1];
  var form = FormApp.create(nama_form).setDescription(deskripsi_form).setIsQuiz(true);

  var item = form.addTextItem();
  item.setTitle('Nomor Induk Siswa').setRequired(true);

  item = form.addTextItem();
  item.setTitle('Nama Siswa').setRequired(true);

  var choices = [];

  for (var i = 5; i < data.length; i++) {
    if (data[i][0] != "") {
      if (data[i][0] == "-") {
        form.addPageBreakItem();
        continue;
      }

      noChoice = true;
      for (var j = 1; j <= 5; j++) {
        if (data[i][j] != '') {
          noChoice = false;
        }
      }

      if (noChoice == false) {
        item = form.addMultipleChoiceItem();
        item.setTitle(data[i][0]).setRequired(true);

        isCorrectA = false;
        isCorrectB = false;
        isCorrectC = false;
        isCorrectD = false;
        isCorrectE = false;

        switch(data[i][6]) {
          case 'A': isCorrectA = true; break;
          case 'B': isCorrectB = true; break;
          case 'C': isCorrectC = true; break;
          case 'D': isCorrectD = true; break;
          case 'E': isCorrectE = true; break;
        }

        choices = []
        if (data[i][1] != "") {
          choices.push(item.createChoice(data[i][1], isCorrectA));
        }
        if (data[i][2] != "") {
          choices.push(item.createChoice(data[i][2], isCorrectB));
        }
        if (data[i][3] != "") {
          choices.push(item.createChoice(data[i][3], isCorrectC));
        }
        if (data[i][4] != "") {
          choices.push(item.createChoice(data[i][4], isCorrectD));
        }
        if (data[i][5] != "") {
          choices.push(item.createChoice(data[i][5], isCorrectE));
        }

        item.setChoices(choices);
      }
      else {
        item = form.addParagraphTextItem();
        item.setTitle(data[i][0]).setRequired(true);
      }

      var point = parseInt(data[i][7]);
      if (point > 0) {
        item.setPoints(point);
      }
    }
  }

  SpreadsheetApp.getUi().alert("Google Form untuk kuis dengan nama \"" + nama_form + "\" selesai dibuat.\nSilahkan cek di Google Drive Anda.");
}