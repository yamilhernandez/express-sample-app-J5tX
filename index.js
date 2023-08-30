import cors from 'cors';
import csv from 'csv-parser';
import Excel from 'exceljs';
import express from 'express';
import * as fs from 'fs';
import multer from 'multer';
import * as os from 'os';
import path from 'path';
import { fileURLToPath } from 'url';

const port = process.env.PORT || 3000;
const app = express();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const upload = multer({ dest: os.tmpdir() });

let cleanResults = [];

let proveedores = {
	'Dra. Janice Sanchez': {
		id: 1,
		vidas: [
			{ id: 124167815, name: 'RIVERA, EDWIN', status: 'NO PLATINO' },
			{ id: 125143125, name: 'CRUZ RAMOS, ESTHER', status: 'PLATINO' },
			{ id: 125375390, name: 'RIVERA CACERES, MYRIAM', status: 'PLATINO' },
			{ id: 125376904, name: 'COLON MERCADO, NELSON E', status: 'PLATINO' },
			{ id: 125330097, name: 'ROSADO, MARIA M', status: 'PLATINO' },
			{ id: 125358337, name: 'CALERO SUAREZ, SYLVIA', status: 'PLATINO' },
			{ id: 125362018, name: 'MIRANDA VEGA, ESTEBAN', status: 'PLATINO' },
			{ id: 125369103, name: 'RODRIGUEZ RIQUELME, CARLOS', status: 'PLATINO' },
			{ id: 125370441, name: 'LEBRON CARDONA, ELBA N', status: 'NO PLATINO' },
		],
	},
	'Dr. Ismael Rosado': {
		id: 2,
		vidas: [
			{ id: 124388216, name: 'VELEZ MERCADO, JUAN', status: 'PLATINO' },
			{ id: 124425343, name: 'PONS RODRIGU, ISABEL', status: 'PLATINO' },
			{ id: 124524212, name: 'NARVAEZ HERNANDEZ, JORGE A', status: 'PLATINO' },
			{ id: 125177657, name: 'TORRES FELICIANO, A', status: 'NO PLATINO' },
			{ id: 125240923, name: 'ROSA-RIVERA, ANIBAL', status: 'NO PLATINO' },
			{ id: 125375795, name: 'SOTO, LUZ N', status: 'PLATINO' },
			{ id: 125323137, name: 'COLON ALMODOVAR, EVELYN', status: 'PLATINO' },
		],
	},
	'Dr. Gustavo Cedeño': {
		id: 3,
		vidas: [
			{ id: 125149528, name: 'MARTINEZ GONZALEZ, LUZ M', status: 'PLATINO' },
			{ id: 125183153, name: 'MAYSONET, CARLOS', status: 'NO PLATINO' },
			{ id: 125288147, name: 'ORTIZ RESTO, LUIS A', status: 'NO PLATINO' },
			{ id: 125359497, name: 'APONTE, DAVID', status: 'NO PLATINO' },
			{ id: 125364530, name: 'OSTOLAZA BURGOS, MARITZA Y', status: 'PLATINO' },
			{ id: 125375882, name: 'BERMUDEZ, PEDRO', status: 'NO PLATINO' },
			{ id: 124118075, name: 'VAZQUEZ-PABO, JESUS M', status: 'PLATINO' },
		],
	},
};

app.use(function (req, res, next) {
	res.header('Access-Control-Allow-Origin', '*'); // update to match the domain you will make the request from
	res.header(
		'Access-Control-Allow-Headers',
		'Origin, X-Requested-With, Content-Type, Accept'
	);
	next();
});

app.use(
	cors({
		origin: '*',
		methods: ['GET'],
	})
);

app.get('/json', (req, res) => {
	res.json({ 'Choo Choo': 'Welcome to your Express app 🚅' });
});

app.get('/test', (req, res) => {
	res.send('Hello World!');
});

app.get('/', (req, res) => {
	res.send('Hello World!');
});

app.get('/download', (req, res) => {
	//const file = `${__dirname}/output/output.xlsx`;
	const file = `/output/output.xlsx`;
	res.download(file); // Set disposition and send it.
});

app.post('/upload', upload.single('file'), function (req, res) {
	const title = req.body.title;
	const file = req.file;

	console.log(title);
	console.log(file);

	fs.createReadStream(file.path)
		.pipe(csv())
		.on('data', (data) => {
			//console.log(data);
			if (data.ToPayAmount != null && moneyToString(data.ToPayAmount) > 0) {
				let ret = moneyToString(data.ToPayAmount) * 0.1;
				let final =
					moneyToString(data.ToPayAmount) - ret - moneyToString(data.Discount);
				if (data.ServiceCode === 'S0250') {
					final = final / 2;
				}

				/* 			gTotal += moneyToString(data.ToPayAmount);
			gDiscount += moneyToString(data.Discount);
			gPercent += roundToTwo(ret);
			gFinal += final;
			console.log(`
      id: ${Object.values(data)[0]}
      Vida: ${data.MemberName} \n
      Total Pagado: ${data.ToPayAmount}
      Retencion 10%: ${roundToTwo(ret)}
      Discount: ${data.Discount}
      Final: ${final}
      `); */
				cleanResults.push({
					id: Object.values(data)[0],
					vida: data.MemberName,
					date: data.srvStartDate,
					total: moneyToString(data.ToPayAmount),
					percent: roundToTwo(ret),
					discount: moneyToString(data.Discount),
					final: final,
				});
			}
		})
		.on('end', () => {
			/* 		console.log(`
  Resultados Finales 
  Total Pagado: ${gTotal}
  Total Retenido: ${gPercent}
  Total Discount: ${gDiscount}
  Total Final: ${gFinal}
  `); */
			sheet();
		});

	res.sendStatus(200);
});

app.listen(port, () => {
	console.log(`Example app listening on port ${port}`);
});

const sheet = async () => {
	const fileName = '/output/output.xlsx';
	//console.log(cleanResults);

	const wb = new Excel.Workbook();
	for (const key in proveedores) {
		let total = 0;
		let discount = 0;
		let percent = 0;
		let final = 0;
		const ws = wb.addWorksheet(key);
		let row = 3;
		ws.getCell('A1').value = 'Nombre';
		ws.getCell('B1').value = 'Fecha';
		ws.getCell('C1').value = 'Total Pagado';
		ws.getCell('D1').value = 'Retencion 10%';
		ws.getCell('E1').value = 'Discount';
		ws.getCell('F1').value = 'Final';
		cleanResults.forEach((element) => {
			//console.log(element);
			proveedores[key].vidas.forEach((vida) => {
				//console.log(vida.id);
				if (element.id == vida.id) {
					ws.getRow(row++).values = [
						element.vida,
						element.date,
						element.total,
						element.percent,
						element.discount,
						element.final,
					];
					total += element.total;
					discount += element.discount;
					percent += element.percent;
					final += element.final;

					cleanResults = cleanResults.filter((item) => item.id !== element.id);
				}
			});
		});
		ws.getCell('H1').value = 'Total';
		ws.getCell('H2').value = total;
		ws.getCell('I1').value = 'total 10%';
		ws.getCell('I2').value = percent;
		ws.getCell('J1').value = 'total Discount';
		ws.getCell('J2').value = discount;
		ws.getCell('K1').value = 'total final';
		ws.getCell('K2').value = final;
	}
	const ws = wb.addWorksheet('Extra Sheet');
	let row = 3;
	ws.getCell('A1').value = 'Nombre';
	ws.getCell('B1').value = 'Fecha';
	ws.getCell('C1').value = 'Total Pagado';
	ws.getCell('D1').value = 'Retencion 10%';
	ws.getCell('E1').value = 'Discount';
	ws.getCell('F1').value = 'Final';
	let eTotal = 0;
	let eDiscount = 0;
	let ePercent = 0;
	let eFinal = 0;
	cleanResults.forEach((element) => {
		ws.getRow(row++).values = [
			element.vida,
			element.date,
			element.total,
			element.percent,
			element.discount,
			element.final,
		];
		eTotal += element.total;
		eDiscount += element.discount;
		ePercent += element.percent;
		eFinal += element.final;

		//cleanResults = cleanResults.filter((item) => item.id !== vida.id);
	});

	ws.getCell('H1').value = 'Total';
	ws.getCell('H2').value = eTotal;
	ws.getCell('I1').value = 'total 10%';
	ws.getCell('I2').value = ePercent;
	ws.getCell('J1').value = 'total Discount';
	ws.getCell('J2').value = eDiscount;
	ws.getCell('K1').value = 'total final';
	ws.getCell('K2').value = eFinal;

	wb.xlsx
		.writeFile(fileName)
		.then(() => {
			console.log('file created');
		})
		.catch((err) => {
			console.log(err.message);
			console.log('error file not created');
		});
};

const moneyToString = (str) => {
	return Number(str.replace(/[^0-9.-]+/g, ''));
};
function roundToTwo(num) {
	return +(Math.round(num + 'e+2') + 'e-2');
}
