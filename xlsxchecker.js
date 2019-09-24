// Compare list items in 2 Excel files
// independent of order
// detect added/removed lines
// detect added/removed subproperties (enumeration)

const XLSX = require("js-xlsx")

const debug = false

const myArgs = process.argv.slice(2)

if (myArgs.length < 2) {
	console.log("FeLIX' XLSX Tag Checker")
	console.log("Missing command line arguments!")
	console.log("Usage: xlsxchecker filename1 filename2")
	process.exit(1)
}

const file1 = myArgs[0]
const file2 = myArgs[1]

compare(file1, file2, "Templates", "Template Tracking", "B", "F")
compare(file1, file2, "Capabilities", "Capability Tracking - Panels", "A", "C")
compare(file1, file2, "Capabilities", "Capability Tracking - Storyboards", "A", "D")

function compare(file1, file2, title, sheet, col1, col2) {
	let dataset = {}
	// Get data from first file
	{
		let workbook = XLSX.readFile(file1);
		let worksheet = workbook.Sheets[sheet];
		
		
		let end = false
		let row = 2
		while (!end) {
			let desired_cell = worksheet[col1 + row];
			let desired_value = (desired_cell ? desired_cell.v : undefined);
		
			let desired_cell2 = worksheet[col2 + row];
			let desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
			
			if (desired_value == undefined) 
				end = true
			else if (desired_value) {
				if (dataset[desired_value]) console.log("WARN ALREADY EXISTING", desired_value)
				dataset[desired_value] = { x:"REMOVED", y: new Set(), deltas: [] }
				let v = desired_value2.replace(/ /g,"").split(",")
				v.map(x => { if (x) dataset[desired_value].y.add(x)})
			}
			row++
		}
	}

	// compare with data from second file
	{
		let workbook = XLSX.readFile(file2);
		let worksheet = workbook.Sheets[sheet];
		
		let end = false
		let row = 2
		while (!end) {
			let desired_cell = worksheet[col1 + row];
			let desired_value = (desired_cell ? desired_cell.v : undefined);
		
			let desired_cell2 = worksheet[col2 + row];
			let desired_value2 = (desired_cell2 ? desired_cell2.v : undefined);
		
			if (debug) console.log(desired_value)
			
			if (desired_value == undefined) 
				end = true
			else if (desired_value) {
				if (dataset[desired_value]) {
					dataset[desired_value].x = "MODIFIED"
		
					let v = desired_value2.replace(/ /g,"").split(",")
					let o = Array.from(	dataset[desired_value].y)
		
					v.map(x => { if (x) {
						if (!o.includes(x)) {
							if (debug) console.log("added", x);
							dataset[desired_value].deltas.push("+" + x)
						}
					} 
					})
		
					o.map(x => { if (x) {
						if (!v.includes(x)) {
							if (debug) console.log("removed", x);
							dataset[desired_value].deltas.push("-" + x)
						}
					}	
					})
					
					if (dataset[desired_value].deltas.length == 0)
						dataset[desired_value].x = "IDENTICAL"
				}
				else
					dataset[desired_value] = { x : "ADDED", y: desired_value2, deltas: [] }
			}
			row++
		}
	}
	//console.log(dataset)

	if (debug) console.log("################################")

	console.log("\n" + title + " Comparison\n")

	// output comparison results
	for(let i in dataset) {
		if (dataset[i].x != "IDENTICAL")
			console.log(i, dataset[i].x, dataset[i].deltas)
	}

	console.log("\n")
}
