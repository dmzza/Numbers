JsOsaDAS1.001.00bplist00�Vscript_�app = Application.currentApplication()
app.includeStandardAdditions = true
app.strictPropertyScope = true

function displayError(errorMessage, buttons) {
	app.displayDialog(errorMessage, { buttons: buttons })
}

Numbers = Application('Numbers')
document = Numbers.documents[0]
try {
	document.activeSheet()
} catch {
	document.open()
	displayError("Spreadsheet not open")
}

const sheet = document.activeSheet()
const table = sheet.tables()[0]

var index = table.rows.byName("Posted Transactions").address() - 1

while (index >= 2) {
	table.rows[index].remove()
	index -= 1
}

table.cells.byName("B2").value = "Memo"
table.cells.byName("D2").value = "Payee"
table.cells.byName("E2").value = "Outflow"
table.cells.byName("F2").value = "Inflow"

table.rows[0].remove()
table.columns[6].remove()
table.columns[2].remove()

path = Path("Users/davidsolis/Desktop/" + table.name() + ".csv")

document.export({"to": path, "as":"CSV"})                              � jscr  ��ޭ