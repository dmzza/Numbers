JsOsaDAS1.001.00bplist00�Vscript_ �
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

function parseNumbers(values) {
	var arr = new Array(values.length)
	for(var i = 0; i < values.length; i++) {
		arr[i] = parseNumber(values[i])
		if(arr[i] instanceof Error) {
			throw arr[i]
		}
	}
	return arr
}


var firstRow = 2
var footerRow = 8
var totalAvailableRow = 1
var numberOfLoans = 6
var constraintsCol = 2
var loanPrincipalsCol = 2
var interestRateCol = 3
var lengthCol = 4
var totalPaidCol = 5
var monthlyPaymentCol = 6
var minimumLengthCol = 9
var maximumLengthCol = 12
var searchSpaceCol = 13
var loanPrincipals = parseNumbers(getValues(getCells(loanCellRange, firstRow, loanPrincipalsCol, numberOfLoans)));
var interestRates = parseNumbers(getValues(getCells(loanCellRange, firstRow, interestRateCol, numberOfLoans)));
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

Progress.totalUnitCount = totalSearchSpace / 1000
Progress.description = "Searching for more optimal loan lengths"
for(var j = 0; j <= totalSearchSpace / 1000; j++) {
  if(j%10000 == 0) {
    Progress.completedUnitCount = j
  }
  // oldLoanLengths = currentLoanLengths.slice();
  currentLoanLengths = nextPermutation(currentLoanLengths, minimumLengths, maximumLengths);
  // setValues(loanLengthsRange, currentLoanLengths, oldLoanLengths);
  totalPaidAmount = totalPaid(loanPrincipals, interestRates, currentLoanLengths)
  totalMonthlyPaymentAmount = totalMonthlyPayment(loanPrincipals, interestRates, currentLoanLengths)
  if(totalPaidAmount < lowestPossibleTotal && totalMonthlyPaymentAmount <= totalAvailable) {
  	oldLoanLengths = optimalPermutation.slice();
    lowestPossibleTotal = totalPaidAmount
    optimalPermutation = currentLoanLengths.slice();
	setValues(loanLengthsRange, optimalPermutation, oldLoanLengths);
    lowestPossibleRange[0].value = lowestPossibleTotal;
    optimalMonthlyPaymentRange[0].value = totalMonthlyPaymentAmount;
  }
}

// setValues(loanLengthsRange, optimalPermutation, currentLoanLengths);

function quit() {
	// setValues(loanLengthsRange, optimalPermutation, currentLoanLengths);
	return true
}


function nextPermutation(currentPermutation, minimumValues, maximumValues) {
  var next = currentPermutation.slice(); // to copy the array instead of pointing to it
  for(var dimension = currentPermutation.length - 1; dimension >= 0; dimension--) {
    if(next[dimension] == minimumValues[dimension]) {
      next[dimension] = maximumValues[dimension];
    } else {
      var jumpDown = Math.ceil(currentPermutation[dimension] / 13.0)
      next[dimension] = Math.max(currentPermutation[dimension] - jumpDown, minimumValues[dimension]);
      break;
    }
  }
  return next;
}

function totalMonthlyPayment(loanPrincipals, interestRates, loanLengths) {
	var monthlyPayments = new Array(loanPrincipals.length)
	for(var i = 0; i < loanPrincipals.length; i++) {
		monthlyPayments[i] = -1 * typeSafePMT(interestRates[i] / 12, loanLengths[i], loanPrincipals[i], 0, 0)
	}
	return monthlyPayments.reduce(function(total, current) { return total + current })
}

function totalPaid(loanPrincipals, interestRates, loanLengths) {
	var totalPaids = new Array(loanPrincipals.length)
	startingPeriod = 1
	dueAtEndOfPeriod = 0
	for(var i = 0; i < loanPrincipals.length; i++) {
		totalPaids[i] = -1 * typeSafeCUMIPMT(interestRates[i] / 12, loanLengths[i], loanPrincipals[i], startingPeriod, loanLengths[i], dueAtEndOfPeriod) + loanPrincipals[i]
	}
	return totalPaids.reduce(function(total, current) { return total + current })
}

/**
* Everything below taken from https://github.com/formulajs/formulajs
*/

var error = {
	nil: new Error('#NULL!'),
	div0: new Error('#DIV/0!'),
	value: new Error('#VALUE!'),
	ref: new Error('#REF!'),
	name: new Error('#NAME?'),
	num: new Error('#NUM!'),
	na: new Error('#N/A'),
	error: new Error('#ERROR!'),
	data: new Error('#GETTING_DATA')
}

function parseNumber(string) {
  if (string === undefined || string === '') {
    return error.value;
  }
  if (!isNaN(string)) {
    return parseFloat(string);
  }

  return error.value;
}
function anyIsError() {
  var n = arguments.length;
  while (n--) {
    if (arguments[n] instanceof Error) {
      return true;
    }
  }
  return false;
}

function CUMIPMT(rate, periods, value, start, end, type) {
  rate = parseNumber(rate);
  periods = parseNumber(periods);
  value = parseNumber(value);
  if (anyIsError(rate, periods, value)) {
    return error.value;
  }
  return typeSafeCUMIPMT(rate, periods, value, start, end, type);
};

function typeSafeCUMIPMT(rate, periods, value, start, end, type) {
  if (rate <= 0 || periods <= 0 || value <= 0) {
    return error.num;
  }

  if (start < 1 || end < 1 || start > end) {
    return error.num;
  }

  if (type !== 0 && type !== 1) {
    return error.num;
  }

  var payment = typeSafePMT(rate, periods, value, 0, type);
  var interest = 0;

  if (start === 1) {
    if (type === 0) {
      interest = -value;
    }
    start++;
  }

  for (var i = start; i <= end; i++) {
    if (type === 1) {
      interest += typeSafeFV(rate, i - 2, payment, value, 1) - payment;
    } else {
      interest += typeSafeFV(rate, i - 1, payment, value, 0);
    }
  }
  interest *= rate;

  return interest;
};

function FV(rate, periods, payment, value, type) {
  // Credits: algorithm inspired by Apache OpenOffice

  value = value || 0;
  type = type || 0;

  rate = parseNumber(rate);
  periods = parseNumber(periods);
  payment = parseNumber(payment);
  value = parseNumber(value);
  type = parseNumber(type);
  if (anyIsError(rate, periods, payment, value, type)) {
    return error.value;
  }
  return typeSafeFV(rate, period, payment, value, type)
};

// each parameter must be a non-optional number
function typeSafeFV(rate, periods, payment, value, type) {
  // Return future value
  var result;
  if (rate === 0) {
    result = value + payment * periods;
  } else {
    var term = Math.pow(1 + rate, periods);
    if (type === 1) {
      result = value * term + payment * (1 + rate) * (term - 1) / rate;
    } else {
      result = value * term + payment * (term - 1) / rate;
    }
  }
  return -result;
};

function PMT(rate, periods, present, future = 0, type = 0) {
  // Credits: algorithm inspired by Apache OpenOffice

  rate = parseNumber(rate);
  periods = parseNumber(periods);
  present = parseNumber(present);
  future = parseNumber(future);
  type = parseNumber(type);
  if (anyIsError(rate, periods, present, future, type)) {
    return error.value;
  }
  return typeSafePMT(rate, periods, present, future, type);
};

// each parameter must be a non-optional number
function typeSafePMT(rate, periods, present, future, type) {
  // Return payment
  var result;
  if (rate === 0) {
    result = (present + future) / periods;
  } else {
    var term = Math.pow(1 + rate, periods);
    if (type === 1) {
      result = (future * rate / (term - 1) + present * rate / (1 - 1 / term)) / (1 + rate);
    } else {
      result = future * rate / (term - 1) + present * rate / (1 - 1 / term);
    }
  }
  return -result;
};



                              !jscr  ��ޭ