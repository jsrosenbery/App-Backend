// server.js
const express=require('express'),cors=require('cors'),multer=require('multer'),
  Papa=require('papaparse'),XLSX=require('xlsx');
const app=express(); app.use(cors()); app.use(express.json()); const upload=multer();
let scheduleData=[],roomMetadata=[];
app.post('/api/schedule',upload.single('file'),(req,res)=>{
  const pw=req.body.password; if(pw!=='Upload2025')return res.status(401).json({error:'Unauthorized'});
  const parsed=Papa.parse(req.file.buffer.toString(),{header:true,skipEmptyLines:true}); scheduleData=parsed.data;
  res.json({success:true});
});
app.get(['/api/schedule','/api/schedule/:term'],(req,res)=>res.json(scheduleData));
app.post('/api/rooms/metadata',upload.single('file'),(req,res)=>{
  const wb=XLSX.read(req.file.buffer,{type:'buffer'}); const raw=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{range:3});
  roomMetadata=raw.map(r=>({campus:r.Campus,building:r.Building,room:r['Room #'].toString(),type:r.Type,capacity:Number(r['# of Desks in Room'])}));
  res.json({success:true});
});
app.get('/api/rooms/metadata',(req,res)=>res.json(roomMetadata));
app.listen(process.env.PORT||3000);
