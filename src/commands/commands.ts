/* global Office */
/* global Word, Excel, PowerPoint, performance, console */

async function addcitation(ref, type) {
  await Word.run(async (context) => {
  const contentControls = context.document.contentControls;
  contentControls.load("items/text");
  await context.sync();
  
  var Index = 0;
  const matchingControlsForIndex = contentControls.items.filter(cc => !(cc.text.includes(" ")));
  matchingControlsForIndex.forEach(cc => {
    let index = removeFirstAndLastChar(cc.text);
    if (index > Index) {
      Index = index;
    }
  });

  const selection = context.document.getSelection();

  // Clear the current selection
  selection.clear();

  Index++;
  // Insert a content control with the text like "[1]" by default, for example the type is not in below two cases.
  var text = "[" + Index + "]";
  if(type === "IEEE") {
    text = "[" + Index + "]";
  }
  else if(type === "Vancouver") {
    text = "(" + Index + ")";
  }
  
  // Insert a content control with the text lik "[1]"
  const contentControl = selection.insertContentControl();
  contentControl.insertText(text, Word.InsertLocation.replace);
  contentControl.tag = "reference";
  contentControl.title = "Reference Marker";
  contentControl.appearance = "BoundingBox";

  const body = context.document.body;
  body.insertParagraph("", Word.InsertLocation.end);

  // Insert a content control with ref at the end of the document
  const range = body.getRange(Word.RangeLocation.end);
  const contentControlw = range.insertContentControl();
  var refAct = text + "    " + ref;

  contentControlw.insertText(refAct, Word.InsertLocation.replace);
  contentControlw.tag = "referenceEnd";
  contentControlw.title = "Reference Marker at End";
  contentControlw.appearance = "BoundingBox";

  await context.sync();
  return "the citation has been added";
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
      return "An error occurred while adding the citation" + "Debug info: " + JSON.stringify(error.debugInfo);
    }
  });
}

// Add a citation with a highlight effect, and remove the highlight after 1 second.
async function addcitationhighlightcontenttemporarily(ref, type) {
  await Word.run(async (context) => {
  const contentControls = context.document.contentControls;
  contentControls.load("items/text");
  await context.sync();
  
  var Index = 0;
  const matchingControlsForIndex = contentControls.items.filter(cc => !(cc.text.includes(" ")));
  matchingControlsForIndex.forEach(cc => {
    let index = removeFirstAndLastChar(cc.text);
    if (index > Index) {
      Index = index;
    }
  });

  const selection = context.document.getSelection();

  // Clear the current selection
  selection.clear();

  Index++;
  // Insert a content control with the text like "[1]" by default, for example the type is not in below two cases.
  var text = "[" + Index + "]";
  if(type === "IEEE") {
    text = "[" + Index + "]";
  }
  else if(type === "Vancouver") {
    text = "(" + Index + ")";
  }
  
  // Insert a content control with the text lik "[1]"
  selection.insertText(text, Word.InsertLocation.replace);
  const contentControl = selection.insertContentControl();
  contentControl.tag = "reference";
  contentControl.title = "Reference Marker";
  contentControl.appearance = "BoundingBox";
  contentControl.load("font/highlightColor");
  contentControl.font.highlightColor = "yellow";

  const body = context.document.body;
  var refAct = text + "    " + ref;
  const para = body.insertParagraph(refAct, Word.InsertLocation.end);

  // Insert a content control with ref at the end of the document
  const contentControlw = para.insertContentControl();
  contentControlw.tag = "referenceEnd";
  contentControlw.title = "Reference Marker at End";
  contentControlw.appearance = "BoundingBox";
  contentControlw.load("font/highlightColor");
  contentControlw.font.highlightColor = "yellow";

  await context.sync();
  const originalColor = contentControl.font.highlightColor;
  const originalColorw = contentControl.font.highlightColor;

  // Wait for 1 second
  await new Promise((resolve) => setTimeout(resolve, 1000));

  // Restore original highlight color
  contentControl.font.highlightColor = originalColor;
  contentControlw.font.highlightColor = originalColorw;
  await context.sync();
  return "the citation has been added";
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
      return "An error occurred while adding the citation" + "Debug info: " + JSON.stringify(error.debugInfo);
    }
  });
}

function getSubstringBeforeFirstSpace(text) {
  const index = text.indexOf(" ");
  return index === -1 ? text : text.substring(0, index);
}

function removeFirstAndLastChar(text) {
  if (text.length <= 2) {
    return ""; // Not enough characters to keep
  }
  return text.substring(1, text.length - 1);
}

function replaceNumberInString(str, newNumber) {
  const start= str[0];
  const end = str[str.length - 1];
  return ""+ start + newNumber + end;
}

async function updatecitation(index, ref, type) {
    if(type !== "IEEE" && type !== "Vancouver") {
      console.log(`Invalid citation type: "${type}"`);
      return "The refence type is not supported" + type;
    }

    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items/text,title,tag");
  
      await context.sync();
  
      const matchingControls = contentControls.items.filter(cc => cc.text.includes(" "));
  
      if (matchingControls.length === 0) {
        console.log(`No citation found containing: "${ref}"`);
        return "No citation found containing: " + ref;
      } else {
        matchingControls.forEach(cc => {
          let indexInternal = removeFirstAndLastChar(getSubstringBeforeFirstSpace(cc.text));
          let newText = "";
          if (indexInternal === index) {
            if(type === "IEEE") {
              newText = '[' + index + ']' + "    " + ref;
            }
            else if(type === "Vancouver") {
              newText = '(' + index + ')' + "    " + ref;
            }
            cc.insertText(newText, Word.InsertLocation.replace);
          }
        });
      }
  
      const matchingControlsForIndex = contentControls.items.filter(cc => !(cc.text.includes(" ")));
      matchingControlsForIndex.forEach(cc => {
        let indexInternal = removeFirstAndLastChar(cc.text);
        let newText = "";
        if (indexInternal === index) {
          if(type === "IEEE") {
            newText = '[' + index + ']';
          }
          else if(type === "Vancouver") {
            newText = '(' + index + ')';
          }
          cc.insertText(newText, Word.InsertLocation.replace);
        }
      });

      await context.sync();
      return "The reference type has been updated to " + type;
    }).catch((error) => {
      console.error("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
        return "An error occurred while removing the citation" + "Debug info: " + JSON.stringify(error.debugInfo);
      }
    });
}

async function removecitation(index) {
  let indexNumber = parseInt(index);
  if(indexNumber <= 0) {
    console.log(`Invalid citation index: "${indexNumber}"`);
    return "The refence index is not valid" + indexNumber;
  }

  return await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items/text,title,tag");

    await context.sync();

    const matchingControls = contentControls.items.filter(cc => cc.text.includes(" "));

    if (matchingControls.length === 0) {
      console.log(`No citation found in the document`);
      return "No citation found in the document";
    } else {
      matchingControls.forEach(cc => {
        let indexCitation = removeFirstAndLastChar(getSubstringBeforeFirstSpace(cc.text));
        if(indexCitation === index) {
          const paragraph = cc.paragraphs.getFirst();
          cc.delete(false);
          paragraph.delete();
        }
        else if(indexCitation > index) {
          let newIndex = indexCitation - 1;
          let newText = replaceNumberInString(getSubstringBeforeFirstSpace(cc.text), newIndex) + "    " + cc.text.substring(cc.text.indexOf("    ") + 4);
          cc.insertText(newText, Word.InsertLocation.replace);
        }
      });
    }

    const matchingControlsForIndex = contentControls.items.filter(cc => !(cc.text.includes(" ")));
    matchingControlsForIndex.forEach(cc => {
      let indexCitation = removeFirstAndLastChar(cc.text);
      if(indexCitation === index) {
        cc.delete(false);
      }
      else if(indexCitation > index) {
        let newIndex = indexCitation - 1;
        let newText = replaceNumberInString(cc.text, newIndex);
        cc.insertText(newText, Word.InsertLocation.replace);
      }
    });

    await context.sync();
    return `the citation at "${index}" has been removed`;
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
      return "An error occurred while removing the citation" + "Debug info: " + JSON.stringify(error.debugInfo);
    }
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
      Office.actions.associate("addcitation", async (message) => {
      const start = performance.now();
      const { Reference: ref, Type: type } = JSON.parse(message);
      //await addcitation(ref, type);
      await addcitationhighlightcontenttemporarily(ref, type);
      const duration = performance.now() - start;
      const result = `Add citation action! completed in ${duration.toFixed(0)} ms.`;
      console.log(`Returning result: "${result}"`);
      return result;
    });

    Office.actions.associate("updatecitation", async (message) => {
      const start = performance.now();
      const {Index: index, Reference: ref, Type: type } = JSON.parse(message);
      const exeresult = await updatecitation(index, ref, type);
      const duration = performance.now() - start;
      const result = exeresult + `. Update citation action! completed in ${duration.toFixed(0)} ms.`;
      console.log(`Returning result: "${result}"`);
      return result;
    });

    Office.actions.associate("removecitation", async (message) => {
      const start = performance.now();
      const { Index: index } = JSON.parse(message);
      const exeresult = await removecitation(index);
      const duration = performance.now() - start;
      const result = exeresult + `. Remove citation action! completed in ${duration.toFixed(0)} ms.`;
      console.log(`Returning result: "${result}"`);
      return result;
    });
  }
});
