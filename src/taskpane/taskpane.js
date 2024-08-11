Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    $("#insertPlaceholder").on("click", insertPlaceholder);
    $("#updatePlaceholder").on("click", updatePlaceholder);    
    $("#placeholderDropdown").on("change", loadPlaceholderContent);
    populateDropdown();
    window.eventListenersAdded = true;
  }
});

async function insertPlaceholder() {
  console.log("Insert placeholder function called");
  const placeholderName = $("#placeholderName").val();
  if (placeholderName) {
    try {
      await Word.run(async context => {
        console.log("Inserting placeholder...");
        const range = context.document.getSelection();
        context.load(range, 'text');
        await context.sync();

        const existingText = range.text;
        if (existingText.includes(`[${placeholderName}]`)) {
          console.log("Platzhalter bereits vorhanden.");
          return;
        }

        range.insertText(`[${placeholderName}]`, Word.InsertLocation.replace);

        const contentControl = range.insertContentControl();
        contentControl.tag = placeholderName;
        contentControl.title = placeholderName;

        await context.sync();
        console.log('Insert and sync completed');
        
        // Aktualisiere das Dropdown nach dem Einfügen
        await populateDropdown();
      });
    } catch (error) {
      console.error("Error inserting placeholder:", error);
    }
  } else {
    alert("Bitte geben Sie einen Platzhalternamen ein.");
  }
}





async function populateDropdown(selectedTag) {
  try {
    await Word.run(async context => {
      const contentControls = context.document.contentControls;
      context.load(contentControls, 'items');
      await context.sync();

      const $dropdown = $("#placeholderDropdown");
      $dropdown.empty();
      contentControls.items.forEach(control => {
        const $option = $("<option></option>").val(control.tag).text(control.tag);
        $dropdown.append($option);
      });

      // Set the newly added placeholder as selected
      if (selectedTag) {
        $dropdown.val(selectedTag);
      }
    });
  } catch (error) {
    console.error("Error populating dropdown:", error);
  }
}


async function loadPlaceholderContent() {
  const selectedPlaceholder = $("#placeholderDropdown").val();
  if (selectedPlaceholder) {
    try {
      await Word.run(async context => {
        const contentControls = context.document.contentControls;
        context.load(contentControls, 'items');
        await context.sync();

        const control = contentControls.items.find(c => c.tag === selectedPlaceholder);
        if (control) {
          const range = control.range;
          range.load('text');
          await context.sync();
          const content = range.text;
          $("#placeholderContent").val(content);
        }
      });
    } catch (error) {
      console.error("Error loading placeholder content:", error);
    }
  }
}

async function updatePlaceholder() {
  const selectedPlaceholder = $("#placeholderDropdown").val();
  const newContent = $("#placeholderContent").val();
  
  if (selectedPlaceholder && newContent) {
    try {
      await Word.run(async context => {
        // Load all content controls
        const contentControls = context.document.contentControls;
        context.load(contentControls, 'items');
        await context.sync();
        
        // Find and update all content controls with the same tag
        contentControls.items.forEach(control => {
          if (control.tag === selectedPlaceholder) {
            control.insertText(newContent, Word.InsertLocation.replace);
          }
        });
        await context.sync();
      });
    } catch (error) {
      console.error("Error updating placeholders:", error);
      alert("Es ist ein Fehler beim Aktualisieren der Platzhalter aufgetreten.");
    }
  } else {
    alert("Bitte wählen Sie einen Platzhalter aus und geben Sie neuen Inhalt ein.");
  }
}

