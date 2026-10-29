/* RB NUTRITION – fabrique le tableur "Suivi prise de masse" dans le navigateur (ExcelJS) */
async function rbTableur() {
  const NOIR='FF0B0B0C', ROUGE='FFE3182D', GRIS='FF9A9AA2', BLANC='FFFFFFFF', INP='FFFFF4C2', BLEU='FF0000FF', VERT='FF0A7A2F';
  const wb = new ExcelJS.Workbook(); wb.creator = 'RB NUTRITION'; wb.calcProperties.fullCalcOnLoad = true;
  const fnt = (o={}) => Object.assign({name:'Arial', size:10}, o);
  const fill = c => ({type:'pattern', pattern:'solid', fgColor:{argb:c}});
  const th = {style:'thin', color:{argb:'FFD0D0D0'}}, box = {top:th, left:th, bottom:th, right:th};
  const center = {horizontal:'center', vertical:'middle', wrapText:true}, left = {horizontal:'left', vertical:'middle', wrapText:true};
  const F = f => ({formula: f});
  const D = (y,m,d) => new Date(Date.UTC(y,m-1,d));

  function sheet(name, tab) { const ws = wb.addWorksheet(name, {views:[{showGridLines:false}], properties:{tabColor:{argb:tab||ROUGE}}}); return ws; }
  function title(ws, t, sub, n) {
    ws.mergeCells(1,1,1,n); let c = ws.getCell(1,1); c.value=t; c.font=fnt({bold:true, italic:true, size:18, color:{argb:BLANC}}); c.fill=fill(NOIR); c.alignment={vertical:'middle', indent:1};
    ws.getRow(1).height=34;
    ws.mergeCells(2,1,2,n); c = ws.getCell(2,1); c.value=sub; c.font=fnt({color:{argb:BLANC}}); c.fill=fill(ROUGE); c.alignment={vertical:'middle', indent:1, wrapText:true};
    ws.getRow(2).height=30;
  }
  function header(ws, row, labels, start=1) {
    labels.forEach((l,i) => { const c = ws.getCell(row, start+i); c.value=l; c.font=fnt({bold:true, color:{argb:BLANC}}); c.fill=fill(NOIR); c.alignment=center; c.border=box; });
    ws.getRow(row).height=32;
  }
  const widths = (ws, arr) => arr.forEach((w,i) => ws.getColumn(i+1).width = w);
  const note = (ws, addr, txt) => { const c = ws.getCell(addr); c.value = txt; c.font = fnt({size:8, color:{argb:GRIS}}); };
  const inputCell = (c, fmt) => { c.fill=fill(INP); c.border=box; c.alignment=center; c.font=fnt({bold:true, color:{argb:BLEU}}); if (fmt) c.numFmt=fmt; };

  /* ---------- Mode d'emploi ---------- */
  let ws = sheet("Mode d'emploi", NOIR);
  title(ws, 'RB NUTRITION · SUIVI PRISE DE MASSE', 'Ouvrier le jour, à la salle le soir. Remplis les cases jaunes, le reste se calcule tout seul.', 6);
  widths(ws, [4,26,22,22,22,22]);
  const lignes = [
    ['1. Mon profil', "Entre ton sexe, âge, taille, poids et activité. Tu obtiens tes calories et macros du jour."],
    ['2. Poids', "Pèse-toi le matin à jeun, après les toilettes. Note chaque pesée : c'est la moyenne de la semaine qui compte, pas un jour isolé."],
    ['3. Repas', "Chaque soir, note tes calories et protéines du jour (appli de comptage ou étiquettes)."],
    ['4. Charges', "Note tes séances : exercice, charge, répétitions. Le tableur repère tes records."],
    ['5. Mensurations', "Une fois par mois, même heure, mètre ruban : bras, poitrine, tour de taille, cuisses."],
    ['Code couleur', "Cases jaunes = à remplir. Cases blanches = calculées, n'y touche pas."],
    ["La règle d'or", "Si ta moyenne de poids ne monte pas pendant 2 semaines, ajoute 150 à 200 kcal par jour."],
    ['Important', "Ces calculs sont des estimations (formule de Mifflin-St Jeor). Ils ne remplacent pas l'avis d'un médecin ou d'un diététicien."]
  ];
  lignes.forEach(([a,b], i) => {
    const r = 4+i; const ca = ws.getCell(r,2); ca.value=a; ca.font=fnt({bold:true, size:11, color:{argb:ROUGE}});
    ws.mergeCells(r,3,r,6); const cb = ws.getCell(r,3); cb.value=b; cb.font=fnt(); cb.alignment=left; ws.getRow(r).height=36;
  });
  ws.getCell(13,2).value = 'Exemple de case à remplir :'; ws.getCell(13,2).font = fnt({color:{argb:GRIS}});
  ws.getCell(13,3).value = 75; inputCell(ws.getCell(13,3));
  note(ws, 'B15', 'rbnutrition.github.io/liens · Code -10 % VitaStrong : RBN76 (partenariat)');

  /* ---------- Mon profil ---------- */
  const p = sheet('Mon profil');
  title(p, 'MON PROFIL', 'Tes besoins du jour pour prendre du muscle sans prendre trop de gras.', 5);
  widths(p, [34,18,4,40,10]);
  header(p, 4, ['À REMPLIR','VALEUR']);
  const ins = [['Sexe (Homme / Femme)','Homme'],['Âge (ans)',33],['Taille (cm)',178],['Poids actuel (kg)',75],
               ["Niveau d'activité",'Modéré (3-4 séances)'],['Surplus calorique (kcal/jour)',300],['Date de début', D(2026,11,2)]];
  ins.forEach(([l,v], i) => { const r = 5+i; const a = p.getCell(r,1); a.value=l; a.font=fnt(); a.border=box; const c = p.getCell(r,2); c.value=v; inputCell(c); });
  p.getCell('B11').numFmt = 'dd/mm/yyyy';
  p.getCell('B5').dataValidation = {type:'list', allowBlank:false, formulae:['"Homme,Femme"']};
  header(p, 4, ["NIVEAUX D'ACTIVITÉ (coefficient)", 'Coef.'], 4);
  const acts = [['Sédentaire (peu de sport)',1.2],['Léger (1-2 séances)',1.375],['Modéré (3-4 séances)',1.55],['Actif (5-6 séances)',1.725],['Très actif (sport + travail physique)',1.9]];
  acts.forEach(([l,v], i) => { const a = p.getCell(5+i,4); a.value=l; a.font=fnt(); a.border=box; const c = p.getCell(5+i,5); c.value=v; c.font=fnt(); c.border=box; c.alignment=center; });
  p.getCell('B9').dataValidation = {type:'list', allowBlank:false, formulae:['$D$5:$D$9']};
  p.getCell('B10').note = '200 à 400 kcal de surplus = prise de masse propre. 300 est un bon départ.';
  note(p, 'D11', "Coefficients : facteurs d'activité classiques (Harris-Benedict / FAO).");
  header(p, 14, ['TES RÉSULTATS','PAR JOUR']);
  const res = [
    ['Métabolisme de base (kcal)', 'ROUND(10*B8+6.25*B7-5*B6+IF(B5="Femme",-161,5),0)', '#,##0'],
    ['Dépense totale (kcal)', 'ROUND(B15*INDEX($E$5:$E$9,MATCH(B9,$D$5:$D$9,0)),0)', '#,##0'],
    ['🎯 CALORIES À MANGER', 'B16+B10', '#,##0'],
    ['🥩 Protéines (g) · 2 g/kg', 'ROUND(2*B8,0)', '0'],
    ['🥑 Lipides (g) · 25 % des calories', 'ROUND(B17*0.25/9,0)', '0'],
    ['🍚 Glucides (g) · le reste', 'ROUND((B17-B18*4-B19*9)/4,0)', '0'],
    ['Prise de poids visée / semaine (kg)', 'ROUND(B8*0.0025,2)&" à "&ROUND(B8*0.005,2)', '@']
  ];
  res.forEach(([l,f,nf], i) => {
    const r = 15+i; const a = p.getCell(r,1); a.value=l; a.font=fnt({bold: i===2||i===3}); a.border=box;
    const c = p.getCell(r,2); c.value=F(f); c.numFmt=nf; c.border=box; c.alignment=center; c.font=fnt({bold:true, size: i===2?13:11, color:{argb: i===2?ROUGE:NOIR}});
  });
  note(p, 'D15', "Formule de Mifflin-St Jeor (estimation la plus fiable chez l'adulte).");
  note(p, 'D18', 'Repère scientifique : 1,6 à 2,2 g/kg pour prendre du muscle (ISSN).');
  note(p, 'D21', 'Viser +0,25 à 0,5 % du poids par semaine limite la prise de gras.');

  /* ---------- Poids ---------- */
  const w = sheet('Poids');
  title(w, 'SUIVI DU POIDS', 'Pèse-toi le matin à jeun. Regarde la MOYENNE de la semaine, pas le chiffre du jour.', 12);
  header(w, 4, ['Semaine','Début','Lun','Mar','Mer','Jeu','Ven','Sam','Dim','Moyenne (kg)','Variation (kg)','Verdict']);
  widths(w, [9,12,8,8,8,8,8,8,8,12,12,22,34]);
  for (let k=0; k<26; k++) {
    const r = 5+k;
    w.getCell(r,1).value = k+1;
    w.getCell(r,2).value = F(`'Mon profil'!$B$11+7*(A${r}-1)`); w.getCell(r,2).numFmt = 'dd/mm';
    for (let c=3; c<=9; c++) { const x = w.getCell(r,c); x.fill=fill(INP); x.numFmt='0.0'; }
    w.getCell(r,10).value = F(`IF(COUNT(C${r}:I${r})=0,"",ROUND(AVERAGE(C${r}:I${r}),2))`); w.getCell(r,10).numFmt='0.00';
    if (k===0) { w.getCell(r,12).value = F('IF(J5="","","Semaine de départ")'); }
    else {
      w.getCell(r,11).value = F(`IF(OR(J${r}="",J${r-1}=""),"",J${r}-J${r-1})`); w.getCell(r,11).numFmt='+0.00;-0.00;0.00';
      w.getCell(r,12).value = F(`IF(K${r}="","",IF(K${r}<'Mon profil'!$B$8*0.0025,"Trop lent : +150 kcal",IF(K${r}>'Mon profil'!$B$8*0.005,"Trop vite : -150 kcal","Parfait, continue")))`);
    }
    for (let c=1; c<=12; c++) { const x = w.getCell(r,c); x.border=box; x.font=fnt({color:{argb:(c>=3&&c<=9)?BLEU:NOIR}}); if (c!==12) x.alignment=center; }
  }
  [[75.2,74.9,75.4,75.1,null,75.0,75.3],[75.4,75.6,75.3,75.7,75.5,null,75.8]].forEach((vals,k) => vals.forEach((v,j) => { if (v!==null) w.getCell(5+k,3+j).value = v; }));
  note(w, 'M5', '← exemple, efface et mets tes pesées');
  w.addConditionalFormatting({ref:'L6:L30', rules:[
    {type:'containsText', operator:'containsText', text:'Parfait', priority:1, style:{font:{color:{argb:VERT}, bold:true}}},
    {type:'containsText', operator:'containsText', text:'Trop', priority:2, style:{font:{color:{argb:ROUGE}, bold:true}}}]});
  w.views = [{showGridLines:false, state:'frozen', xSplit:2, ySplit:4}];

  /* ---------- Repas ---------- */
  const m = sheet('Repas');
  title(m, 'CALORIES & PROTÉINES DU JOUR', 'Note tes totaux chaque soir. Objectif : être dans le vert 6 jours sur 7.', 8);
  header(m, 4, ['Date','Calories mangées','Protéines mangées (g)','Objectif kcal','Objectif protéines','Calories %','Protéines %','Journée']);
  widths(m, [12,14,14,12,13,11,11,16,4,30,12]);
  for (let k=0; k<84; k++) {
    const r = 5+k;
    m.getCell(r,1).value = F(`'Mon profil'!$B$11+${k}`); m.getCell(r,1).numFmt = 'ddd dd/mm';
    [2,3].forEach(c => { const x = m.getCell(r,c); x.fill=fill(INP); x.numFmt='#,##0'; });
    m.getCell(r,4).value = F(`'Mon profil'!$B$17`); m.getCell(r,4).numFmt='#,##0';
    m.getCell(r,5).value = F(`'Mon profil'!$B$18`); m.getCell(r,5).numFmt='0';
    m.getCell(r,6).value = F(`IF(B${r}="","",B${r}/D${r})`); m.getCell(r,6).numFmt='0%';
    m.getCell(r,7).value = F(`IF(C${r}="","",C${r}/E${r})`); m.getCell(r,7).numFmt='0%';
    m.getCell(r,8).value = F(`IF(OR(B${r}="",C${r}=""),"",IF(AND(F${r}>=0.9,G${r}>=0.9),"✅ Validée","❌ À rattraper"))`);
    for (let c=1; c<=8; c++) { const x = m.getCell(r,c); x.border=box; x.alignment=center; x.font=fnt({color:{argb:(c===2||c===3)?BLEU:NOIR}}); }
  }
  m.getCell('B5').value=2950; m.getCell('C5').value=148; m.getCell('B6').value=2600; m.getCell('C6').value=120;
  note(m, 'J5', '← exemple, efface et mets tes chiffres');
  m.getCell('J7').value = 'Jours validés'; m.getCell('J7').font = fnt({bold:true});
  m.getCell('K7').value = F('COUNTIF(H5:H88,"✅*")&" / "&COUNTA(B5:B88)'); m.getCell('K7').font = fnt({bold:true, size:12, color:{argb:ROUGE}});
  m.addConditionalFormatting({ref:'H5:H88', rules:[
    {type:'containsText', operator:'containsText', text:'✅', priority:1, style:{font:{color:{argb:VERT}, bold:true}}},
    {type:'containsText', operator:'containsText', text:'❌', priority:2, style:{font:{color:{argb:ROUGE}, bold:true}}}]});
  m.views = [{showGridLines:false, state:'frozen', xSplit:1, ySplit:4}];

  /* ---------- Charges ---------- */
  const c4 = sheet('Charges');
  title(c4, 'SUIVI DES CHARGES', "Si tu soulèves plus lourd de mois en mois, tu prends du muscle. 1RM estimé = formule d'Epley.", 8);
  header(c4, 4, ['Date','Exercice','Charge (kg)','Répétitions','Séries','1RM estimé (kg)','Record perso','Tonnage (kg)']);
  header(c4, 4, ['EXERCICES','Meilleur 1RM'], 11);
  widths(c4, [12,24,11,12,9,13,14,12,12,4,22,14]);
  const exos = ['Développé couché','Squat','Soulevé de terre','Développé militaire','Rowing barre','Tractions','Dips','Curl biceps','Presse à cuisses','Hip thrust'];
  exos.forEach((e,i) => {
    const r = 5+i; const a = c4.getCell(r,11); a.value=e; a.font=fnt(); a.border=box;
    const b = c4.getCell(r,12); b.value = F(`IF(COUNTIF($B$5:$B$304,K${r})=0,"",_xlfn.MAXIFS($F$5:$F$304,$B$5:$B$304,K${r}))`); b.numFmt='0.0'; b.border=box; b.alignment=center; b.font=fnt({bold:true, color:{argb:ROUGE}});
  });
  note(c4, 'K16', 'Tu peux ajouter tes exercices en bas de la liste (K15).');
  for (let k=0; k<300; k++) {
    const r = 5+k;
    for (let c=1; c<=5; c++) c4.getCell(r,c).fill = fill(INP);
    c4.getCell(r,1).numFmt = 'dd/mm/yy';
    c4.getCell(r,2).dataValidation = {type:'list', allowBlank:true, formulae:['$K$5:$K$15']};
    c4.getCell(r,6).value = F(`IF(OR(C${r}="",D${r}=""),"",IF(D${r}=1,C${r},ROUND(C${r}*(1+D${r}/30),1)))`); c4.getCell(r,6).numFmt='0.0';
    c4.getCell(r,7).value = F(`IF(F${r}="","",IF(F${r}>=_xlfn.MAXIFS($F$5:$F$304,$B$5:$B$304,B${r}),"🏆 Record",""))`);
    c4.getCell(r,8).value = F(`IF(OR(C${r}="",D${r}="",E${r}=""),"",C${r}*D${r}*E${r})`); c4.getCell(r,8).numFmt='#,##0';
    for (let c=1; c<=8; c++) { const x = c4.getCell(r,c); x.border=box; x.alignment=center; x.font=fnt({color:{argb: c<=5?BLEU:NOIR}}); }
  }
  [[D(2026,11,3),'Développé couché',80,8,4],[D(2026,11,3),'Rowing barre',60,10,4],[D(2026,11,10),'Développé couché',82.5,8,4]]
    .forEach((row,k) => row.forEach((v,j) => c4.getCell(5+k,1+j).value = v));
  note(c4, 'I5', '← exemple');
  c4.getCell('G5').font = fnt({bold:true, color:{argb:ROUGE}});
  c4.addConditionalFormatting({ref:'G5:G304', rules:[{type:'containsText', operator:'containsText', text:'Record', priority:1, style:{font:{color:{argb:ROUGE}, bold:true}}}]});
  c4.views = [{showGridLines:false, state:'frozen', xSplit:0, ySplit:4}];

  /* ---------- Mensurations ---------- */
  const s = sheet('Mensurations');
  title(s, 'MENSURATIONS', '1 fois par mois, le matin, mètre ruban. Bras qui monte + tour de taille stable = prise de masse propre.', 8);
  header(s, 4, ['Mois','Date','Poids (kg)','Bras (cm)','Poitrine (cm)','Tour de taille (cm)','Cuisse (cm)','Bras gagné depuis le début (cm)']);
  widths(s, [8,12,11,11,13,14,12,18,4,16]);
  for (let k=0; k<12; k++) {
    const r = 5+k; s.getCell(r,1).value = k+1;
    for (let c=2; c<=7; c++) { const x = s.getCell(r,c); x.fill=fill(INP); x.numFmt='0.0'; }
    s.getCell(r,2).numFmt = 'dd/mm/yy';
    s.getCell(r,8).value = F(`IF(OR(D${r}="",$D$5=""),"",D${r}-$D$5)`); s.getCell(r,8).numFmt='+0.0;-0.0;0.0';
    for (let c=1; c<=8; c++) { const x = s.getCell(r,c); x.border=box; x.alignment=center; x.font=fnt({color:{argb:(c>=2&&c<=7)?BLEU:NOIR}}); }
  }
  [D(2026,11,2),75,36.5,100,82,57].forEach((v,j) => s.getCell(5,2+j).value = v);
  note(s, 'J5', '← exemple');

  return wb.xlsx.writeBuffer();
}

async function rbTelecharger(btn) {
  const txt = btn ? btn.innerHTML : '';
  try {
    if (btn) btn.innerHTML = '<strong>⏳ Préparation du tableur…</strong>';
    if (!window.ExcelJS) await new Promise((ok, ko) => { const sc = document.createElement('script'); sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js'; sc.onload = ok; sc.onerror = ko; document.head.appendChild(sc); });
    const buf = await rbTableur();
    const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'Suivi-prise-de-masse-RB-NUTRITION.xlsx';
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    if (btn) btn.innerHTML = '<strong>✅ Tableur téléchargé !</strong><small>Ouvre-le avec Excel, Google Sheets ou Numbers</small>';
  } catch (e) {
    if (btn) btn.innerHTML = txt;
    alert('Le téléchargement a échoué. Réessaie depuis un autre navigateur.');
  }
}
