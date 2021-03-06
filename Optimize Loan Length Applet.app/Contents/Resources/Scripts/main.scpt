JsOsaDAS1.001.00bplist00�Vscript_�
Numbers = Application('Numbers');
document = Numbers.documents()[0]
sheet = document.activeSheet()
loanTable = sheet.tables()[0]
constraintTable = sheet.tables()[1]
loanCellRange = loanTable.cellRange()
constraintCellRange = constraintTable.cellRange()

function getCells(range, row, column, numberOfRows = 1) {
	return range.cells.whose({ _and: [ 
		{ _match: [ ObjectSpecifier().row.address, { '>=': row } ] },
		{ _match: [ ObjectSpecifier().row.address, { '<': row + numberOfRows } ] },
		{ _match: [ ObjectSpecifier().column.address, column ] }
	]});
}

function getValues(cells) {
	var arr = new Array(cells.length)
	for(var i = 0; i < cells.length; i++) {
		arr[i] = cells[i].value()
	}
	return arr
}

function setValues(cells, values, oldValues) {
	for(var i = 0; i < cells.length; i++) {
		if(values[i] !== oldValues[i]) {
			cells[i].value = values[i]
		}
	}
}


var firstRow = 2
var footerRow = 9
var totalAvailableRow = 1
var numberOfLoans = 7
var constraintsCol = 2
var lengthCol = 4
var totalPaidCol = 5
var monthlyPaymentCol = 6
var minimumLengthCol = 9
var maximumLengthCol = 12
var searchSpaceCol = 13
var loanLengthsRange = getCells(loanCellRange, firstRow, lengthCol, numberOfLoans);
var totalPaidRange = getCells(loanCellRange, footerRow, totalPaidCol);
var lowestPossibleRange = getCells(loanCellRange, footerRow + 1, totalPaidCol);
var totalMonthlyPaymentRange = getCells(loanCellRange, footerRow, monthlyPaymentCol);
var optimalMonthlyPaymentRange = getCells(loanCellRange, footerRow + 1, monthlyPaymentCol);
var minimumLengths = getValues(getCells(loanCellRange, firstRow, minimumLengthCol, numberOfLoans));
var maximumLengths = getValues(getCells(loanCellRange, firstRow, maximumLengthCol, numberOfLoans));
var searchSpaces = getValues(getCells(loanCellRange, firstRow, searchSpaceCol, numberOfLoans));
var totalSearchSpace = getCells(loanCellRange, footerRow, searchSpaceCol)[0].value();
var totalAvailable = getCells(constraintCellRange, totalAvailableRow, constraintsCol)[0].value();


var lowestPossibleTotal = totalPaidRange[0].value();
var currentLoanLengths = getValues(loanLengthsRange);
var optimalPermutation = currentLoanLengths.slice();

for(var j = 0; j <= totalSearchSpace / 38055555; j++) {
  oldLoanLengths = currentLoanLengths.slice();
  currentLoanLengths = nextPermutation(currentLoanLengths, minimumLengths, maximumLengths);
  setValues(loanLengthsRange, currentLoanLengths, oldLoanLengths);
  if(totalPaidRange[0].value() < lowestPossibleTotal && totalMonthlyPaymentRange[0].value() <= totalAvailable) {
    lowestPossibleTotal = totalPaidRange[0].value()
    optimalPermutation = currentLoanLengths.slice();
    lowestPossibleRange[0].value = lowestPossibleTotal;
    optimalMonthlyPaymentRange[0].value = totalMonthlyPaymentRange[0].value();
  }
}

setValues(loanLengthsRange, optimalPermutation, currentLoanLengths);

function quit() {
	setValues(loanLengthsRange, optimalPermutation, currentLoanLengths);
	return true
}


function nextPermutation(currentPermutation, minimumValues, maximumValues) {
  var next = currentPermutation.slice(); // to copy the array instead of pointing to it
  for(var dimension = currentPermutation.length - 1; dimension >= 0; dimension--) {
    if(next[dimension] == minimumValues[dimension]) {
      next[dimension] = maximumValues[dimension];
    } else {
      var jumpDown = Math.ceil(currentPermutation[dimension] / 3.0)
      next[dimension] = Math.max(currentPermutation[dimension] - jumpDown, minimumValues[dimension]);
      break;
    }
  }
  return next;
}
                              � jscr  ��ޭ