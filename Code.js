function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Prepare")
    .addItem("Write script", "createScript")
    // .addItem("Create video", "createVideo")
    .addToUi();
}

function createScript() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Templates");
  const rankSheet = ss.getSheetByName("Ranks");
  const miscSheet = ss.getSheetByName("Misc");
  const eventsSheet = ss.getSheetByName("Events");
  const mbSheet = ss.getSheetByName("Merit Badges");
  const instructionSheet = ss.getSheetByName("Instructions");
  const positionSheet = ss.getSheetByName("Positions");
  const summarySheet = ss.getSheetByName("Summary");

  function getTemplate(title) {
    return templateSheet
      .getRange("A1:B")
      .getDisplayValues()
      .find((row) => row[0] === title)[1];
  }

  const opening = getTemplate("Opening");
  const recapTemp = getTemplate("Event Recap");
  const transToAdvance = getTemplate("Transition");
  const totinTemp = getTemplate("Totin Chip");
  const firemnTemp = getTemplate("Firem'n Chit");
  const closing = getTemplate("Closing");
  const startRankTemp = getTemplate("Each Rank");
  const scoutTemp = getTemplate("Scout");
  const tenderfootTemp = getTemplate("Tenderfoot");
  const secondClassTemp = getTemplate("Second Class");
  const firstClassTemp = getTemplate("First Class");
  const starTemp = getTemplate("Star");
  const lifeTemp = getTemplate("Life");
  const eagleTemp = getTemplate("Eagle");
  const endRankTemp = getTemplate("End Each Rank");
  const recruiterTemp = getTemplate("Recruiter");
  const eachRecruiterTemp = getTemplate("Each Recruiter");
  const mbTemp = getTemplate("MBs");
  const positionTransition = getTemplate("Position Transition");
  const splTemp = getTemplate("SPL");
  const plTemp = getTemplate("PLs");
  const aplTemp = getTemplate("APL");
  const positionTemp = getTemplate("Other Positions");
  const patrolCupTemp = getTemplate("Patrol Cup");
  const summaryTemp = getTemplate("Summary");

  const year = instructionSheet.getRange("E2").getDisplayValue();
  const month = instructionSheet.getRange("E4").getDisplayValue();
  const splB = instructionSheet.getRange("E6").getDisplayValue();
  const splG = instructionSheet.getRange("E7").getDisplayValue();

  let doc;
  let header;
  let newDoc = instructionSheet.getRange("C8").isChecked();

  if (!newDoc) {
    try {
      doc = DocumentApp.openByUrl(
        instructionSheet.getRange("B11").getDisplayValue()
      );
      header = doc.getHeader().clear();
      if (!header) header = doc.addHeader();
    } catch {
      let ui = SpreadsheetApp.getUi();
      let response = ui.alert(
        "Invalid URL. Would you like to create a new document?",
        ui.ButtonSet.YES_NO
      );

      // Process the user's response.
      if (response == ui.Button.YES) {
        newDoc = true;
      } else {
        return;
      }
    }
  }

  if (newDoc) {
    doc = DocumentApp.create(`T182 ${month} ${year} CoH Script`);
    header = doc.addHeader();
  }

  header
    .setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 })
    .appendParagraph(`T182 ${month} ${year} CoH Script`)
    .setAttributes({
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]:
        DocumentApp.HorizontalAlignment.CENTER,
      [DocumentApp.Attribute.FONT_FAMILY]: "Times New Roman",
      [DocumentApp.Attribute.FONT_SIZE]: 24,
      [DocumentApp.Attribute.BOLD]: true,
    });
  header
    .appendParagraph(" ")
    .setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 })
    .appendHorizontalRule();
  header
    .appendParagraph(" ")
    .setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 12 });

  const style = {
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]:
      DocumentApp.HorizontalAlignment.JUSTIFY,
    [DocumentApp.Attribute.FONT_FAMILY]: "Times New Roman",
    [DocumentApp.Attribute.FONT_SIZE]: 12,
    [DocumentApp.Attribute.LINE_SPACING]: 2,
    [DocumentApp.Attribute.INDENT_FIRST_LINE]: 25,
    [DocumentApp.Attribute.ITALIC]: false,
  };

  const script = doc.getBody().setAttributes(style).clear();

  function addParagraph(content) {
    return script.appendParagraph(content).setAttributes(style);
  }

  addParagraph(
    opening.replace("{year}", year).replace("{month}", month) + "\n"
  );

  let events = recapTemp + "\n";
  const allEvents = eventsSheet.getRange("A2:Z").getDisplayValues();
  for (let i = 0; i < allEvents.length; i++) {
    let event = allEvents[i];
    if (event[0] === "") break;

    events += event[0];
    //patches
    if (event[1] !== "") {
      let attendees = event[1];
      for (let j = 2; j < event.length; j++) {
        if (event[j] === "") break;

        attendees += ", " + event[j];
      }
      events += ": " + attendees;
    }

    events += "\n";
  }
  addParagraph(events);
  addParagraph(transToAdvance);
  addParagraph();

  function totinFiremn(template, data, type) {
    let scouts = data[0];
    for (let i = 1; i < data.length; i++) {
      let scout = data[i][0];
      if (scout === "") break;

      scouts += `, ${scout}`;
    }

    if (scouts !== "") {
      addParagraph(template.replace(`{${type}Scouts}`, scouts) + "\n");
    }
  }

  const allTotin = miscSheet.getRange("A2:A").getDisplayValues();
  if (allTotin[0][0]) totinFiremn(totinTemp, allTotin, "totin");

  const allFiremn = miscSheet.getRange("B2:B").getDisplayValues();
  if (allFiremn[0][0]) totinFiremn(firemnTemp, allFiremn, "firemn");

  const allRecruiters = miscSheet.getRange("C2:D").getDisplayValues();
  let recruiters = "";
  for (let i = 0; i < allRecruiters.length; i++) {
    let recruiterPair = allRecruiters[i];
    if (recruiterPair[0] === "") break;

    recruiters += `\n${eachRecruiterTemp
      .replace("{recruited}", recruiterPair[1])
      .replace("{recruiter}", recruiterPair[0])}`;
  }
  if (recruiters) addParagraph(recruiterTemp + recruiters + "\n");

  const allMBs = mbSheet
    .getRange("A2:Z")
    .getDisplayValues()
    .filter((e) => e[0] !== "");

  if (allMBs.length > 0) {
    addParagraph(mbTemp);
  }

  script.appendParagraph(`${splB} Presents:`).setAttributes({
    ...style,
    [DocumentApp.Attribute.ITALIC]: true,
    [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
  });

  for (let i = 0; i < allMBs.length; i++) {
    const mb = allMBs[i];
    if (mb[0] === "-") {
      script.appendParagraph(`${splG} Presents:`).setAttributes({
        ...style,
        [DocumentApp.Attribute.ITALIC]: true,
        [DocumentApp.Attribute.INDENT_FIRST_LINE]: 0,
      });
    } else {
      addParagraph(
        `${mb[0]}: ${mb
          .filter((e) => e !== "")
          .slice(1)
          .join(", ")}`
      );
    }
  }

  addParagraph();

  const allRanks = rankSheet.getRange("A2:B").getDisplayValues();

  function rank(template, rank) {
    const data = allRanks.filter((item) => item[1] === rank).map((d) => d[0]);
    if (data.length === 0) return;

    addParagraph(
      startRankTemp.replace("{rank}", rank).replace("{scouts}", data.join(", "))
    );
    addParagraph(template);
    addParagraph(endRankTemp);
    addParagraph();
  }

  rank(scoutTemp, "Scout");
  rank(tenderfootTemp, "Tenderfoot");
  rank(secondClassTemp, "Second Class");
  rank(firstClassTemp, "First Class");
  rank(starTemp, "Star");
  rank(lifeTemp, "Life");
  rank(eagleTemp, "Eagle");

  if (month === "June") {
    addParagraph(`${positionTransition}\n`);

    for (let i = 0; i < 2; i++) {
      const positions = positionSheet
        .getRange(`${i === 0 ? "A" : "C"}3:${i === 0 ? "B" : "D"}`)
        .getDisplayValues();

      const spl = positions.filter((p) => p[1] === "SPL")[0][0];
      const pls = positions.filter((p) => p[1] === "PL").map((p) => p[0]);
      const apls = positions.filter((p) => p[1] === "APL").map((p) => p[0]);
      const otherPositions = positions.filter(
        (p) => !["SPL", "PL", "APL", ""].includes(p[1])
      );

      addParagraph(`${[splB, splG][i]}: ${splTemp.replace("{spl}", spl)}`);

      addParagraph(
        `${plTemp
          .replace("{spl}", spl)
          .replace(
            "{pls}",
            pls.length === 1
              ? pls[0]
              : [
                  pls.slice(0, pls.length - 1).join(", "),
                  pls[pls.length - 1],
                ].join(pls.length === 2 ? " & " : ", & ")
          )}`
      );

      for (let j = 0; j < pls.length; j++) {
        addParagraph(
          `${aplTemp
            .replace("{spl}", spl)
            .replace("{pl}", pls[j])
            .replace("{apl}", apls[j])}`
        );
      }

      if (otherPositions.length > 0) {
        addParagraph(
          `${spl}:\n${otherPositions
            .map((pos) =>
              positionTemp
                .replace("{position}", pos[1])
                .replace("{scout}", pos[0])
            )
            .join("\n")}`
        );
      }

      addParagraph();
    }

    addParagraph(patrolCupTemp.replace("{year}", year));
    addParagraph();
    addParagraph(
      summaryTemp
        .replace("{newMembers}", summarySheet.getRange("B1").getDisplayValue())
        .replace("{totalMbs}", summarySheet.getRange("B2").getDisplayValue())
        .replace("{totalRanks}", summarySheet.getRange("B3").getDisplayValue())
    );
    addParagraph();
  }

  addParagraph(closing);

  SpreadsheetApp.getUi().alert(`Script created:\n${doc.getUrl()}`);
}

// function createVideo() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   const slideshow = SlidesApp.openById("1HAEbYKA3P5LeOUt8Fg8aOr7ohWoO_u0KqJeBoZBrjWk");
//   const slides = slideshow.getSlides();

//   const photos = DriveApp.getFolderById("1BrXIMlwj8v0fxiKoR-j1i-v-GnIrx5jo").getFiles();

//   // while (slides.length > 7) {
//   //   slides[slides.length - 1].remove();
//   // }

//   while (photos.hasNext()){
//     let photo = photos.next();
//     console.log(photo.getName());

//     let newSlide = slides[Math.floor(Math.random() * 7)].duplicate();
//     newSlide.move(slides.length + 1);
//     newSlide.setSkipped(false);

//     newSlide.insertImage(photo.getDownloadUrl()).alignOnPage(SlidesApp.AlignmentPosition.CENTER);
//   }

//   SpreadsheetApp.getUi().alert(`Video created:\n${slideshow.getUrl()}`);
// }
