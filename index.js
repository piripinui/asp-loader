const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
    sourceFile: 'data/asp-upload.xlsx'
});

console.log(JSON.stringify(result, null, 3));

let aspDefs = [];
let aspLocDefs = [];

const createASPDef = (aspRow) => {
	let anASP = {
		name: aspRow['A'],
		description: aspRow['B'],
		parentPosition: aspRow['C'],
		status: aspRow['D'],
		spatialType: aspRow['E'],
		class: aspRow['F'],
		positionFilledBy: aspRow['G'],
		operatingStatus: aspRow['K'],
		inService: aspRow['L'] == 'TRUE' ? true : false,
		networkAnalysisEnabled: aspRow['M'] == 'TRUE' ? true : false,
		location: aspRow['I'],
		equipmentComponentNumber: aspRow['X'],
		linearFrom: Number(aspRow['S']),
		linearTo: Number(aspRow['T']),
		linearLength: Number(aspRow['U']),
		linearUOM: aspRow['V'],
		legalEntity: aspRow['P'],
		costCenter: aspRow['Q'],
		dependentCustomer: aspRow['N'],
		modelAuthority: aspRow['O'],
		account: aspRow['R'],
		catalogueItem: aspRow['W']
	};
	
	anASP = removeUndefinedsAndNulls(anASP);
	
	return anASP;
}

const createASPLocDef = (aspRow) => {
	console.log(aspRow['D'].replace(/\"/g, ''));//.split(";")[0].split(","));
	let geom, geojsonType, lumadaGeom, geomType = aspRow['C'];
	switch(geomType) {
		case 'Point': {
			geojsonType = 'Point';
			geom = [
				Number(aspRow['D'].replace(/\"/g, '').split(";")[0].split(",")[0]),
				Number(aspRow['D'].replace(/\"/g, '').split(";")[0].split(",")[1])
			];
			lumadaGeom = [
				{
					x: Number(aspRow['D'].replace(/\"/g, '').split(";")[0].split(",")[0]),
					y: Number(aspRow['D'].replace(/\"/g, '').split(";")[0].split(",")[1])
				}
			]
			break;
		}
		case 'Linear': {
			geojsonType = 'LineString';
			geom = [];
			lumadaGeom = [];
			let parsedCoords = aspRow['D'].replace(/\"/g, '').split(";")[0].split(",");
			for (let i = 0; i < parsedCoords.length - 1; i += 2) {
				geom.push([parsedCoords[i], parsedCoords[i + 1]]);
				lumadaGeom.push({
					x: parsedCoords[i],
					y: parsedCoords[i + 1]
				});
			}
			break;
		}
		case 'Area': {
			geojsonType = 'Polygon';
			geom = [];
			lumadaGeom = [];
			let parsedCoords = aspRow['D'].replace(/\"/g, '').split(";")[0].split(",");
			for (let i = 0; i < parsedCoords.length - 1; i += 2) {
				geom.push([parsedCoords[i], parsedCoords[1]]);
				lumadaGeom.push({
					x: parsedCoords[i],
					y: parsedCoords[i + 1]
				});
			}
			break;
		}
	}		
			
	let anASPLoc = {
		locationId: aspRow['A'],
		locationDescription: aspRow['B'],
		locationGIS: {
			type: geojsonType,
			geometry: lumadaGeom
		}
	}

	anASPLoc = removeUndefinedsAndNulls(anASPLoc);
		
	return anASPLoc;
}

const removeUndefinedsAndNulls = (anObj) => {
	// Remove any 'undefined' or null fields.
	for (const propertyName in anObj) {
		if (typeof anObj[propertyName] == 'undefined' || anObj[propertyName] === null)
			delete anObj[propertyName];
	}
	
	return anObj;
}

for (const sheetName in result) {
	switch(sheetName) {
		case 'Asset System Positions': {
			let asps = result[sheetName];
			console.log(asps);
			
			for (let i = 1; i < asps.length; i++) {
				// Row 0 is the header row.
				let anASP = createASPDef(asps[i]);
				aspDefs.push(anASP);
			}
			break;
		}
		case 'Lists': {
			console.log("Sheet " + sheetName + " is ignored.");
			break;
		}
		case 'ASP Locations': {
			let aspLocs = result[sheetName];
			
			for (let i = 1; i < aspLocs.length; i++) {
				// Row 0 is the header row.
				let anASPLoc = createASPLocDef(aspLocs[i]);
				aspLocDefs.push(anASPLoc);
			}
			break;
		}
		default: {
			console.log("Sheet name " + sheetName + " not recognised.");
		}
	}
		
}

let amqpInput = {
	data: [
		{
			locations: aspLocDefs
		},
		{
			assetsystempositions: aspDefs
		}
	],
	sourceSystem: 'uploader'
}

console.log(JSON.stringify(amqpInput,null, 3));