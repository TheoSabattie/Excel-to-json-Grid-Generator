var Cell 	  = require("./cell");
const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];



function ExcelSheetHelper(sheet){
	this.sheet = sheet;
	this.cache = {}
}



ExcelSheetHelper.prototype.getCellAt = function (x, y){
	xCache = this.cache[x] = this.cache[x] || {};

	if (xCache[y]){
		return xCache[y];
	}

	var cellId = this.getColumnIDFromX(x) + (y + 1);

	return xCache[y] = new Cell(
		x, 
		y, 
		cellId,
		this.sheet[cellId]
	);
};



ExcelSheetHelper.prototype.getColumnIDFromX = function (x){
	var lettersAmount   = letters.length;
	var lettersIDAmount = Math.floor(x/lettersAmount)
	var ID              = "";

	for (var i = 0; i < lettersIDAmount; i++){
		ID += letters[0];
	}

	ID += letters[x % lettersAmount];

	return ID;
};



ExcelSheetHelper.prototype.getUpCellNeighbourOfCell = function (cell){
	return this.getCellAt(cell.x, cell.y - 1);
};


ExcelSheetHelper.prototype.getDownCellNeighbourOfCell = function (cell){
	return this.getCellAt(cell.x, cell.y + 1);
};


ExcelSheetHelper.prototype.getLeftCellNeighbourOfCell = function (cell){
	return this.getCellAt(cell.x - 1, cell.y);
};


ExcelSheetHelper.prototype.getRightCellNeighbourOfCell = function (cell){
	return this.getCellAt(cell.x + 1, cell.y);
};


ExcelSheetHelper.prototype.hasContentInUpCellNeighbourOfCell = function (cell){
	return Boolean(this.getCellAt(cell.x, cell.y - 1).content);
};


ExcelSheetHelper.prototype.hasContentInDownCellNeighbourOfCell = function (cell){
	return Boolean(this.getCellAt(cell.x, cell.y + 1).content);
};


ExcelSheetHelper.prototype.hasContentInLeftCellNeighbourOfCell = function (cell){
	return Boolean(this.getCellAt(cell.x - 1, cell.y).content);
};


ExcelSheetHelper.prototype.hasContentInRightCellNeighbourOfCell = function (cell){
	return Boolean(this.getCellAt(cell.x + 1, cell.y).content);
};


ExcelSheetHelper.prototype.getNeighboursWithContent = function (cell){
	var neighbours = [];

	neighbours.push(this.getUpCellNeighbourOfCell(cell));
	neighbours.push(this.getDownCellNeighbourOfCell(cell));
	neighbours.push(this.getRightCellNeighbourOfCell(cell));
	neighbours.push(this.getLeftCellNeighbourOfCell(cell));

	for (var i = neighbours.length -1; i > -1; i--){
		if (!neighbours[i].content){
			neighbours.splice(i, 1);
		}
	}

	return neighbours;
};

ExcelSheetHelper.letters = letters.slice(0);

module.exports = ExcelSheetHelper;