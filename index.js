var fs                = require('fs-extra');
var GridJsonGenerator = require("./gridjsongenerator")
const EXCEL_EXT       = ".xlsx";

init();


function init(){
    fs.readJson("./gridConfig.json", function(err, config){
        if (err){
            if (err.code == 'ENOENT'){
                console.log("ConfigException : Config does not exist. Creation of default config file 'gridConfig.json'. Please, complete it! Try to use it...");
                config = getDefaultConfig();
		writeDefaultConfig(config);
            } else {
		console.log("ConfigException : " + err);
	        return;
            }
        }
	    
        var errorMsg = getConfigurationErrorMessage(config);

        if (errorMsg){
            throw errorMsg.bgRed;
            return;
        }

        var gridJsonGenerator = new GridJsonGenerator(config);

        fs.writeJson('grid.json', gridJsonGenerator.generatedJson, function(err){
	    if (err){ 
                console.log("WriteJsonException : " + err);
            } else {
                console.log("grid.json is write with success!");
	    }
        });
    });
}



function writeDefaultConfig(config){
    fs.writeJson('./gridConfig.json', config, function (err) {   
        if (err) {
            console.log(err)
	}
    });
}



function getDefaultConfig(){
    return {
	filePath        : "./grid" + EXCEL_EXT,
        cellsSheetName  : "Cells",
        paramsSheetName : "Params"
    };
}



function getConfigurationErrorMessage(config){
    if (!config.filePath){
        return "ConfigException : filePath is not defined";
    } else if (!objIsString(config.filePath)){
    	return "ConfigException : filePath must be a string";
    } else if (config.filePath.indexOf(EXCEL_EXT) == -1){
        return "ConfigException : filePath is not an excel file (need '" + EXCEL_EXT + "' extension)";
    }

    if (!config.cellsSheetName){
        return "ConfigException : cellsSheetName is not defined)";
    } else if (!objIsString(config.cellsSheetName)){
    	return "ConfigException : cellsSheetName must be a string";
    }

    if (!config.paramsSheetName){
        return "ConfigException : paramsSheetName is not defined";
    } else if (!objIsString(config.paramsSheetName)){
    	return "ConfigException : paramsSheetName must be a string";
    }

    return false;
}



function objIsString(obj){
    return typeof(obj) == "string";
}
