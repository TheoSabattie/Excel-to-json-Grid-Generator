var XLSX              = require("xlsx");
var ExcelSheetHelper  = require("./excelSheetHelper");
var colors            = require("colors");


function GridJsonGenerator(config){
    this.generatedJson = this.generateJsonFromConfig(config);
}



GridJsonGenerator.prototype.generateJsonFromConfig = function (config){
    var xlsx = XLSX.readFile(config.filePath);

    if (!xlsx.Sheets[config.cellsSheetName]){
        throw ("ConfigException : CellsSheetName define on configuration ('" + config.cellsSheetName + "') does not exist on xlsx file").bgRed;
    } else if (!xlsx.Sheets[config.paramsSheetName]){
        throw ("ConfigException : ParamsSheetName define on configuration ('" + config.paramsSheetName + "') does not exist on xlsx file").bgRed;
    }

    var cellsSheet  = new ExcelSheetHelper(xlsx.Sheets[config.cellsSheetName]);
    var paramsSheet = new ExcelSheetHelper(xlsx.Sheets[config.paramsSheetName]);

    var json      = this.getParamsFromSheet(paramsSheet); 
    json["cells"] = this.getCellsFromSheet(cellsSheet);

    return json;
}



GridJsonGenerator.prototype.getParamsFromSheet = function(sheet){
    var params = {};
    var cell;
    var i = 0;

    while (true){
        cell = sheet.getCellAt(0, i);

        if (!cell.content){
            break;
        }

        params[cell.getValue()] = this.getValueParamFromSheetAndKeyCell(sheet, cell);
        i++;
    }

    return params;
}



GridJsonGenerator.prototype.getValueParamFromSheetAndKeyCell = function(sheet, cell){
    var paramValues = [];
    var paramName   = cell.getValue();
    var cell        = sheet.getRightCellNeighbourOfCell(cell);
    var letter;

    paramValues.push(cell.getValue());

    while (true){
        cell = sheet.getRightCellNeighbourOfCell(cell);

        if (!cell.content){
            break;
        }

        paramValues.push(cell.getValue());
    }

    if (!paramValues.length){
        return null;
    }

    if (paramName == "givenLetters"){
        for (var i in paramValues){
            letter = paramValues[i];
            letter = (typeof(value) == "string")? letter.toUpperCase() : letter;

            if (ExcelSheetHelper.letters.indexOf(letter) == -1){
                throw ("InvalidLetterException :: Letter must be alphabetic! Enter '" + letter + "' on param givenLetters").bgRed;
            }
        }
    }


    return (paramValues.length == 1)? paramValues.pop() : paramValues;
}



GridJsonGenerator.prototype.getCellsFromSheet = function(sheet){
    var cell = this.findFirstCellOnFirstLineFromCellSheet(sheet);

    if (!cell){
        throw "XLSXException : There is not filled cell on first line on CellSheet!".bgRed;
    }

    var basicCells = this.getConnectedNeighbour([cell], sheet);
    var cells      = [];
    var value;

    for (var i in basicCells){
        cell  = basicCells[i];
        value = cell.getValue();
        value = (typeof(value) == "string")? value.toUpperCase() : value;
        
        if (ExcelSheetHelper.letters.indexOf(value) == -1){
            throw ("InvalidLetterException :: Letter must be alphabetic! Enter '" + value + "' on cell " + cell.id).bgRed;
        }

        cells.push({
            x      : cell.x,
            letter : value,
            y      : cell.y
        });
    }

    return cells;
}


GridJsonGenerator.prototype.getConnectedNeighbour = function(neighbours, sheet){
    var cell;
    var cell2;
    var newNeighbours;

    for (var i in neighbours){
        cell = neighbours[i];
        newNeighbours = sheet.getNeighboursWithContent(cell);

        for (var j in newNeighbours){
            cell2 = newNeighbours[j];

            if (neighbours.indexOf(cell2) == -1){
                neighbours.push(cell2);
                this.getConnectedNeighbour(neighbours, sheet);
            }
        }
    }

    return neighbours;
} 


GridJsonGenerator.prototype.findFirstCellOnFirstLineFromCellSheet = function(cellsSheet){
    for (var i = 0; i < 100; i++){
        cell = cellsSheet.getCellAt(i, 0);

        if (cell.content){
            return cell;
        }
    }

    return null;
}


module.exports = GridJsonGenerator;