// Prueba de humo del backend con una hoja de cálculo simulada (no necesita Google).
// Uso: node tests/backend.test.js
const fs=require('fs'),vm=require('vm');
const code=fs.readFileSync(require('path').join(__dirname,'..','apps-script','Codigo.gs'),'utf8');
function Sheet(name){this.name=name;this.d=[];this.formulas={};}
Sheet.prototype={
 getName(){return this.name},setName(n){this.name=n;return this},
 get(r,c){const row=this.d[r-1];return row&&row[c-1]!==undefined?row[c-1]:''},
 set(r,c,v){while(this.d.length<r)this.d.push([]);const row=this.d[r-1];while(row.length<c)row.push('');row[c-1]=v},
 getLastRow(){for(let i=this.d.length;i>0;i--)if(this.d[i-1].some(v=>v!==''&&v!==null))return i;return 0},
 getLastColumn(){return Math.max(0,...this.d.map(r=>{for(let i=r.length;i>0;i--)if(r[i-1]!=='')return i;return 0}))},
 getRange(a,b,c,e){if(typeof a==='string')return new Range(this,1,1,1,1);return new Range(this,a,b,c||1,e||1)},
 getDataRange(){return new Range(this,1,1,Math.max(1,this.getLastRow()),Math.max(1,this.getLastColumn()))},
 appendRow(row){const r=this.getLastRow()+1;row.forEach((v,j)=>this.set(r,j+1,v))},
 deleteRow(r){this.d.splice(r-1,1)},deleteRows(r,n){this.d.splice(r-1,n)},
 clear(){this.d=[];this.formulas={}},createTextFinder(t){return new Range(this,1,1,Math.max(1,this.getLastRow()),Math.max(1,this.getLastColumn())).createTextFinder(t)},
 copyTo(ss){const s=new Sheet(this.name+' copy');s.d=JSON.parse(JSON.stringify(this.d));ss._sheets.push(s);return s},
 setFrozenRows(){},setFrozenColumns(){},setRowHeight(){},setColumnWidths(){},setColumnWidth(){}
};
function Range(sh,r,c,nr,nc){Object.assign(this,{sh,r,c,nr,nc})}
const chain=['setFontWeight','setFontColor','setBackground','setWrap','setVerticalAlignment','setNumberFormat','setFontSize','setFontFamily','setFontStyle','setDataValidation'];
chain.forEach(m=>Range.prototype[m]=function(){return this});
Object.assign(Range.prototype,{
 setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>this.sh.set(this.r+i,this.c+j,x)));return this},
 setValue(v){this.sh.set(this.r,this.c,v);return this},
 setFormula(f){this.sh.formulas[this.r+','+this.c]=f;return this},
 getValues(){const o=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++)row.push(this.sh.get(this.r+i,this.c+j));o.push(row)}return o},
 getRow(){return this.r},
 createTextFinder(t){const self=this;let entire=false;return{matchEntireCell(x){entire=x;return this},findNext(){for(let i=0;i<self.nr;i++)for(let j=0;j<self.nc;j++){const v=String(self.sh.get(self.r+i,self.c+j)).toLowerCase(),q=String(t).toLowerCase();if(entire?v===q:v.includes(q))return new Range(self.sh,self.r+i,self.c+j,1,1)}return null}}}
});
const ss={_sheets:[],getSheetByName(n){return this._sheets.find(s=>s.name===n)||null},insertSheet(n,i){const s=new Sheet(n);if(i===0)this._sheets.unshift(s);else this._sheets.push(s);return s},getSheets(){return this._sheets},setSpreadsheetTimeZone(){},setActiveSheet(){}};
const cache={};
const ui={alert:(t,m)=>{console.log('[ALERT]',t,'|',String(m).replace(/\n/g,' / '));return 'YES'},prompt:(t,m)=>({getSelectedButton:()=> 'OK',getResponseText:()=>ctx.__answers.shift()}),ButtonSet:{OK:1,YES_NO:2,OK_CANCEL:3},Button:{YES:'YES',OK:'OK'},createMenu(){const m={addItem(){return m},addSeparator(){return m},addToUi(){}};return m}};
const ctx={SpreadsheetApp:{getActive:()=>ss,getActiveSpreadsheet:()=>ss,getUi:()=>ui,newDataValidation:()=>({requireValueInList(){return this},build(){return{}}})},
 LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},CacheService:{getScriptCache:()=>({get:k=>cache[k]||null,put:(k,v)=>cache[k]=v,remove:k=>delete cache[k]})},
 Session:{getActiveUser:()=>({getEmail:()=>'rafael@unifranz.edu.bo'})},Utilities:{formatDate:(d)=>d.toISOString().slice(0,10)},
 ContentService:{createTextOutput:t=>({t,setMimeType(){return this}}),MimeType:{JSON:1,JAVASCRIPT:2}},Logger:{log:console.log},console,__answers:[]};
// base de RRHH ficticia
const base=ss.insertSheet('Base_de_estudiantes');
base.d=[['sede','modalidad','anho_plan_estudio','carrera','semestre_pertenencia','primer_nombre','segundo_nombre','primer_apellido','segundo_apellido','ci','nacionalidad','correo_institucional','celular'],
 ['CBB','PRESENCIAL',2026,'NGE',1,'JUAN','CARLOS','PEREZ','LOPEZ','1','BOLIVIANA','cbbe.juancarlos.perez.lo@unifranz.edu.bo','7'],
 ['LPZ','PRESENCIAL',2026,'ICO',1,'ANA','','ROJAS','VARGAS','2','BOLIVIANA','LPZE.ana.rojas.va@unifranz.edu.bo ','7'],
 ['SCZ','PRESENCIAL',2026,'GAS',1,'LUIS','','PAZ','SOTO','3','BOLIVIANA','scze.luis.paz.so@unifranz.edu.bo','7'],
 ['SCZ','PRESENCIAL',2017,'ARQ',2,'MARIA','','DIAZ','CRUZ','4','BOLIVIANA','scze.maria.diaz.cr@unifranz.edu.bo','7'],
 ['EAT','PRESENCIAL',2026,'MED',1,'ROSA','','MAMANI','QUISPE','5','BOLIVIANA','eate.rosa.mamani.qu@unifranz.edu.bo','7']];
vm.createContext(ctx);vm.runInContext(code,ctx);
const J=o=>JSON.stringify(o);
ctx.configuracionInicial();
console.log('PADRON rows',ss.getSheetByName('PADRON').getLastRow()-1, J(ss.getSheetByName('PADRON').d.slice(0,3)));
console.log('PADRON L2 formula',ss.getSheetByName('PADRON').formulas['2,12']);
console.log('CARRERAS tail',J(ss.getSheetByName('CARRERAS').d.slice(-2)));
const get=(a,extra)=>JSON.parse(ctx.doGet({parameter:Object.assign({action:a},extra||{})}).t);
console.log('validar ok',J(get('validar',{correo:'cbbe.juancarlos.perez.lo@unifranz.edu.bo'})));
console.log('validar mayus',J(get('validar',{correo:'lpze.ana.rojas.va@unifranz.edu.bo'})).slice(0,120));
console.log('validar plan 2017',J(get('validar',{correo:'scze.maria.diaz.cr@unifranz.edu.bo'})));
console.log('validar nada',J(get('validar',{correo:'x.y@unifranz.edu.bo'})));
// alumno nuevo agregado a la base despues del padron
base.d.push(['CBB','PRESENCIAL',2026,'PSI',1,'NUEVO','','ALUMNO','X','6','B','cbbe.nuevo.alumno.x@unifranz.edu.bo','7']);
console.log('validar nuevo en base',J(get('validar',{correo:'cbbe.nuevo.alumno.x@unifranz.edu.bo'})).slice(0,160),'padron rows',ss.getSheetByName('PADRON').getLastRow()-1, ss.getSheetByName('PADRON').formulas[(ss.getSheetByName('PADRON').getLastRow())+',12']);
const pl={id:'ES-AAAAA-BBBBBB',correo:'cbbe.juancarlos.perez.lo@unifranz.edu.bo',sede:'Cochabamba',sede_cod:'CBB',carrera:'Ingeniería Comercial',carrera_cod:'ICO',datos_corregidos:true,p04:'Sí',p05:4,p06:5,p07:3,p08:2,p09:'Notas',p10:4,p11:5,p12:'=mal',r01_idx:[1,2,3,4],r02:4,r03:5,r04:'',dur_s:90,movil:true,version:'2.0'};
const post=d=>JSON.parse(ctx.doPost({postData:{contents:J({action:'guardar',data:d})}}).t);
console.log('guardar',J(post(pl)));
console.log('guardar dup id',J(post(pl)));
console.log('guardar otro id mismo correo',J(post({...pl,id:'ES-CCCCC-DDDDDD'})));
console.log('validar ya',J(get('validar',{correo:pl.correo})));
const R=ss.getSheetByName('RESPUESTAS'),H=ctx.cabResp_();
H.forEach((h,i)=>{ if(i<14||i>=22&&i<36) console.log(i+1,h.slice(0,40),'=>',J(R.get(2,i+1)).slice(0,50))});
// cerrar
ctx.menuCerrar(); console.log('cerrada',J(get('validar',{correo:'lpze.ana.rojas.va@unifranz.edu.bo'}))); ctx.menuAbrir();
// anular
ctx.__answers=['CBBE.juancarlos.perez.lo@unifranz.edu.bo','Se equivocó de carrera']; ctx.menuAnular();
console.log('resp rows',R.getLastRow()-1,'papelera rows',ss.getSheetByName('PAPELERA').getLastRow()-1);
console.log('validar tras anular',J(get('validar',{correo:pl.correo})).slice(0,60));
// restaurar
ctx.__answers=['es-aaaaa-bbbbbb']; ctx.menuRestaurar(); console.log('resp rows',R.getLastRow()-1);
// excluir
ctx.__answers=['eate.rosa.mamani.qu@unifranz.edu.bo','Retiro']; ctx.menuExcluir(); console.log('validar excluida',J(get('validar',{correo:'eate.rosa.mamani.qu@unifranz.edu.bo'})));
ctx.menuActualizarPadron();
ctx.__answers=['eate.rosa.mamani.qu@unifranz.edu.bo']; ctx.menuReincorporar(); console.log('validar reincorp',J(get('validar',{correo:'eate.rosa.mamani.qu@unifranz.edu.bo'})).slice(0,60));
// agregados manual (plan 2017 no cumple filtro pero agregado)
const ag=ss.getSheetByName('AGREGADOS'); ag.appendRow(['scze.maria.diaz.cr@unifranz.edu.bo','SCZ','PRESENCIAL',2017,'ARQ',2,'MARIA','','DIAZ','CRUZ','caso especial']);
console.log('validar agregado',J(get('validar',{correo:'scze.maria.diaz.cr@unifranz.edu.bo'})).slice(0,80));
// prueba
console.log('prueba',J(post({id:'PRUEBA-XYZ12',correo:'prueba@unifranz.edu.bo',sede:'Cochabamba',sede_cod:'CBB',carrera:'X',carrera_cod:'X',p04:'Sí',r01_idx:[1]})));
// nuevo periodo
ctx.__answers=['I-2027']; ctx.menuNuevoPeriodo(); console.log('sheets',ss._sheets.map(s=>s.name).join(','),'resp rows',R.getLastRow()-1, 'periodo', ctx.config_().periodo);
ctx.menuEstado();
const rs=ss.getSheetByName('RESUMEN'); for(let r=1;r<=14;r++) console.log(r,[1,2,3,4].map(c=>String(rs.get(r,c)).slice(0,90)).join(' || '));
