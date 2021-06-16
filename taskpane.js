Office.onReady(info => {
	document.getElementById('begu').textContent='Taskpanel loaded in '+info.host;
	if (info.host === Office.HostType.Word) {
	}
});

async function run() {
	Office.context.document.customXmlParts.addAsync(
        '<root categoryId="1" xmlns="http://tempuri.org"><item name="Cheap Item" price="$193.95"/><item name="Expensive Item" price="$931.88"/></root>',
        function (result) {});
	return Word.run(async context => {
    	const paragraph = context.document.body.insertParagraph("Hello again!", Word.InsertLocation.end);
    	paragraph.font.color = "blue";
    	await context.sync();
  	});
}
