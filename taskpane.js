Office.onReady(info => {
	document.getElementById('begu').textContent='Taskpanel loaded in '+info.host;
	if (info.host === Office.HostType.Word) {
	}
});

async function run() {
	return Word.run(async context => {
    	const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    	paragraph.font.color = "blue";
    	await context.sync();
  	});
}
