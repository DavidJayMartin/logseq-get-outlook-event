

async function getCurrentBlock() {
  let currentBlock = await logseq.App.getCurrentBlock();
  if (currentBlock) {
    currentBlockId = currentBlock.uuid;  // Store the block ID (UUID) in the variable
    console.log("Current block ID:", currentBlockId);
  } else {
    console.log("No block is currently selected.");
  }
  return currentBlock;
}

async function getCurrentPage(block) {
  const page = await logseq.Editor.getPage(block.page.id);
  if (page) {
    currentPageName = page.name;
    console.log("Page name:", currentPageName);
    
    if (page.journal) {
      console.log("This is a journal page. Date:", page.journal["date"]);
    } else {
      console.log("This is not a journal page.");
    }
  } else {
    console.log("No page is currently open.");
  }

  return page;
}

function journalDayToDate(journalDay) {
  const y = Math.floor(journalDay / 10000);
  const m = Math.floor((journalDay % 10000) / 100) - 1;
  const d = journalDay % 100;
  return new Date(y, m, d);
}





//Insert the activity to the block
async function getActivity (e) {
  
  const currentBlock = await getCurrentBlock();
  const currentPage = await getCurrentPage(currentBlock);
  const pageDate = await journalDayToDate(currentPage?.journalDay);
  console.log(pageDate);
  
  logseq.Editor.insertBlock(e.uuid, `Hello World`, {before: true});

  console.log('Activity Added to the Block');
}

const main = async () => {
  console.log('GetActivity Plugin Loaded');
  
  logseq.Editor.registerSlashCommand('Get Activity', async (e) => {
  getActivity(e)
  })
}

logseq.ready(main).catch(console.error);