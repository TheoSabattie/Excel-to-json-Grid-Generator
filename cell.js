function Cell(x, y, id, content){
	this.x 		 = x;
	this.y 		 = y;
	this.id 	 = id;
	this.content = content;
}

Cell.prototype.getValue = function (){
	if (!this.content){
		return null;
	}

	return this.content.v;
};

module.exports = Cell;