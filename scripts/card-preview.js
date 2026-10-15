// Genera el JSON de la tarjeta de Teams (la que envía POST /api/publish) a
// partir de CSVs de ejemplo, sin tocar ningún webhook real. El archivo
// resultante se puede pegar en https://adaptivecards.io/designer (host app:
// "Microsoft Teams") para validarla visualmente antes de un cambio real.
//
// Uso: npm run card:preview
// Salida: scripts/card-previews/<caso>.json (uno por caso de prueba)
'use strict';

const fs = require('fs');
const path = require('path');
const { processFiles, buildTeamsMessageFitted, MAX_CARD_BYTES } = require('../server.js');

const OUT_DIR = path.join(__dirname, 'card-previews');
fs.mkdirSync(OUT_DIR, { recursive: true });

// URL base de ejemplo solo para armar los links de los botones Action.OpenUrl
// en el preview — no se hace ninguna llamada de red a esta URL.
const BASE_URL = 'https://filters-b3zk.onrender.com';

function file(name, text) {
  return { originalname: name, buffer: Buffer.from(text, 'utf8') };
}

function makeRun(draftCsv, specialtyCsv) {
  const result = processFiles(file('draft.csv', draftCsv), file('specialty.csv', specialtyCsv));
  return {
    id: 'preview-run-id',
    createdAt: new Date().toISOString(),
    result
  };
}

function writeCase(name, draftCsv, specialtyCsv) {
  const run = makeRun(draftCsv, specialtyCsv);
  const message = buildTeamsMessageFitted(run, BASE_URL);
  const json = JSON.stringify(message, null, 2);
  const outPath = path.join(OUT_DIR, `${name}.json`);
  fs.writeFileSync(outPath, json);
  const bytes = Buffer.byteLength(JSON.stringify(message), 'utf8');
  const flag = bytes > MAX_CARD_BYTES ? '  ** OVER MAX_CARD_BYTES — revisar buildTeamsMessageFitted **' : '';
  console.log(`[${name}] -> ${outPath} (${bytes} bytes / limit ${MAX_CARD_BYTES})${flag}`);
}

const HEADER = 'CSN,PACE MRN,First Name,Last Name,Patient Address,Time,Visit Date,Provider/Resource,Type,Appt Notes';

// Caso 1: participante encontrado — aparece en Filters y, al no estar en
// Specialty, también en "Missing in Specialty" (dos "Found in" distintos).
{
  const draft = [
    HEADER,
    '1,SDP100,Maria,Caro,"100 Oak St",09:00,2026-01-05,Provider A,Normal,',
    '2,SDP101,John,Smith,"200 Pine St",10:00,2026-01-05,Provider A,Normal,'
  ].join('\n');
  const specialty = ['CSN', '2'].join('\n'); // CSN 1 (Maria) falta en Specialty
  writeCase('participant-found', draft, specialty);
}

// Caso 2: participante no encontrado en ninguna de las 3 categorías.
{
  const draft = [
    HEADER,
    '1,SDP200,John,Smith,"200 Pine St",10:00,2026-01-05,Provider A,Normal,',
    '2,SDP201,Jane,Doe,"300 Elm St",11:00,2026-01-05,Provider A,Normal,'
  ].join('\n');
  const specialty = ['CSN', '1', '2'].join('\n');
  writeCase('participant-not-found', draft, specialty);
}

// Caso 3: más de 15 filas en una tabla de resultados — ninguna CSN está en
// Specialty, así que las 20 caen en "Missing in Specialty" y deben truncarse.
{
  const rows = [HEADER];
  for (let i = 1; i <= 20; i++) {
    rows.push(`${i},SDP${300 + i},First${i},Last${i},"${i} Test St",09:0${i % 10},2026-01-05,Provider A,Normal,`);
  }
  const draft = rows.join('\n');
  const specialty = ['CSN'].join('\n');
  writeCase('many-rows', draft, specialty);
}
