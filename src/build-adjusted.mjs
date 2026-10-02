import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const projectDir = path.dirname(fileURLToPath(import.meta.url));
const output = path.resolve(projectDir, "../outputs/SAPPT_V1_Diploma_Project_Presentation_Adjusted.pptx");
await fs.mkdir(path.dirname(output), { recursive: true });
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const navy = "#183B70", blue = "#24539A", ink = "#17243A", muted = "#52627A", pale = "#EEF4FB", yellow = "#F2D53C", white = "#FFFFFF";

function box(slide, x, y, w, h, fill, radius = "rounded-xl") {
  const config = { geometry: "roundRect", position: { left:x, top:y, width:w, height:h }, fill, line: { style:"solid", fill, width:0 } };
  if (radius !== "none") config.borderRadius = radius;
  return slide.shapes.add(config);
}
function text(slide, value, x, y, w, h, size=24, color=ink, bold=false, align="left") {
  const s = slide.shapes.add({ geometry:"textbox", position:{left:x,top:y,width:w,height:h}, fill:"none", line:{style:"solid",fill:"none",width:0} });
  s.text = value; s.text.style = { fontSize:size, color, bold, alignment:align, fontFace:"Aptos" }; return s;
}
function top(slide, title, section, n) {
  slide.background.fill = white;
  box(slide, 0, 0, 1280, 18, navy, "none");
  text(slide, section.toUpperCase(), 72, 42, 300, 20, 13, blue, true);
  text(slide, title, 72, 74, 1100, 54, 34, navy, true);
  box(slide, 72, 142, 72, 5, yellow, "none");
  text(slide, `SAPPT · ${String(n).padStart(2,"0")}`, 1040, 46, 168, 18, 12, muted, true, "right");
}
function bulletList(slide, items, x, y, w, size=22, gap=49) {
  items.forEach((item,i) => { box(slide,x,y+i*gap+8,10,10,yellow,"rounded-full"); text(slide,item,x+26,y+i*gap,w-26,34,size,ink,false); });
}
function three(slide, entries) {
  entries.forEach((e,i)=>{ const x=72+i*390; box(slide,x,218,348,280,pale); text(slide,e[0],x+28,248,292,55,24,navy,true); box(slide,x+28,321,46,4,yellow,"none"); text(slide,e[1],x+28,349,292,116,18,muted,false); });
}

// 1
{ const s=deck.slides.add(); s.background.fill=navy; box(s,0,0,1280,18,yellow,"none"); text(s,"DIPLOMA PROJECT PRESENTATION",72,68,380,22,14,yellow,true); text(s,"Student Academic Performance &\nProgress Tracker (SAPPT)",72,118,740,110,42,white,true); text(s,"A web-based application for academic monitoring and improvement",72,250,700,40,24,"#DCE8FB",false); box(s,72,338,188,5,yellow,"none"); text(s,"Track performance. Identify weaknesses. Improve results.",72,370,710,45,23,white,true); text(s,"Presented by\nAnuj Arun Deshmukh · Vishal Balande · Vishal Dake",72,563,620,60,18,"#DCE8FB",false); text(s,"SAPPT V1",1024,586,184,30,22,yellow,true,"right"); }
// 2
{ const s=deck.slides.add(); top(s,"Project overview","Introduction",2); text(s,"SAPPT helps students and administrators bring academic information into one place.",72,190,1080,42,25,ink,false); three(s,[["Academic records","Students can view and manage their academic records in a structured way."],["Performance tracking","Subject-wise performance highlights areas that may need attention."],["Study planning","Schedules, exam dates and a study timer support day-to-day planning."]]); }
// 3
{ const s=deck.slides.add(); top(s,"Problem statement","Need",3); text(s,"Students often need a simple way to monitor progress, plan study time and keep track of exam-related information.",72,194,790,70,27,ink,false); box(s,908,188,300,330,navy); text(s,"SAPPT focuses on a practical student workflow",938,230,240,80,25,white,true); text(s,"Records\n→ Performance\n→ Schedule\n→ Improvement",938,350,230,118,21,"#DCE8FB",true); bulletList(s,["Academic details may be spread across different records.","Weak subjects may not be noticed early.","Exam dates and study planning need a clear view."],72,326,750,20,58); }
// 4
{ const s=deck.slides.add(); top(s,"Project objectives","Objectives",4); three(s,[["Performance analysis","Analyze subject-wise performance to identify weak areas and subjects needing improvement."],["Schedule & exam management","Manage study timetables and track upcoming exam dates so students stay organized."],["Improvement & timer","Provide improvement suggestions and a built-in study timer to support time management."]]); }
// 5
{ const s=deck.slides.add(); top(s,"Scope of the project","Scope",5); text(s,"SAPPT is designed as a web-based academic monitoring application for two core user groups.",72,190,930,38,25,ink,false); box(s,72,270,530,265,pale); text(s,"Students",108,310,220,38,28,navy,true); bulletList(s,["Access academic records","Plan study schedules","Track upcoming exams","Use the study timer"],108,370,420,19,40); box(s,678,270,530,265,navy); text(s,"Administrators",714,310,300,38,28,white,true); bulletList(s,["Add and update student records","Manage schedules","Maintain exam dates"],714,370,400,19,48); }
// 6
{ const s=deck.slides.add(); top(s,"Users and responsibilities","Users",6); box(s,72,202,530,340,pale); text(s,"Student",108,244,300,38,28,navy,true); text(s,"Uses SAPPT to review academic information, track performance, plan study schedules, follow exam dates and use the study timer.",108,312,430,118,22,muted,false); box(s,678,202,530,340,navy); text(s,"Administrator",714,244,330,38,28,white,true); text(s,"Maintains student records and helps keep schedules and exam dates up to date.",714,312,420,96,22,"#DCE8FB",false); }
// 7
{ const s=deck.slides.add(); top(s,"Functional requirements","Requirements",7); three(s,[["Student access","Student registration and login provide access to the application."],["Academic management","Academic records and performance analysis support progress tracking."],["Planning support","Timetable, exam management and the study timer support preparation."]]); }
// 8
{ const s=deck.slides.add(); top(s,"System modules","Modules",8); text(s,"The project is organized around seven practical modules.",72,184,820,38,24,ink,false); const modules=["Student Registration & Login","Academic Record Management","Performance Analysis Module","Improvement Suggestions System","Timetable & Schedule Manager","Exam Management Module","Study Timer Feature"]; modules.forEach((m,i)=>{const col=i<4?0:1; const row=i<4?i:i-4; const x=72+col*570, y=250+row*70; box(s,x,y,42,42,blue); text(s,String(i+1),x,y+8,42,25,18,white,true,"center"); text(s,m,x+62,y+6,420,30,22,ink,true);}); }
// 9
{ const s=deck.slides.add(); top(s,"Non-functional requirements","Quality",9); three(s,[["Performance & speed","Pages should load quickly and process academic data without unnecessary delay."],["Security & reliability","Student records and academic data should be stored securely, with attention to uptime and data loss prevention."],["Usability & design","The interface should be simple, intuitive and easy for students and administrators to use."]]); }
// 10
{ const s=deck.slides.add(); top(s,"Expected output","Outcome",10); text(s,"SAPPT brings the project features together in a single academic-monitoring workflow.",72,188,940,40,25,ink,false); box(s,72,270,1136,180,pale); text(s,"Student records",112,322,230,30,23,navy,true); text(s,"Performance view",390,322,260,30,23,navy,true); text(s,"Schedule & exams",704,322,260,30,23,navy,true); text(s,"Study timer",1002,322,160,30,23,navy,true); text(s,"Organized academic information",112,374,230,28,18,muted); text(s,"Subject-wise tracking",390,374,260,28,18,muted); text(s,"Better preparation",704,374,260,28,18,muted); text(s,"Time management",1002,374,160,28,18,muted); text(s,"The intended result is a clearer view of academic progress and practical support for study planning.",72,518,1080,40,23,ink,false); }
// 11
{ const s=deck.slides.add(); top(s,"Benefits of SAPPT","Advantages",11); three(s,[["Better visibility","Students can review academic information and identify areas that need attention."],["Improved planning","Schedules, exam dates and the study timer support organized preparation."],["Useful administration","Administrators can maintain records, schedules and exam information in one system."]]); }
// 12
{ const s=deck.slides.add(); top(s,"Project flow","Working approach",12); const steps=[["1","Register / log in"],["2","Manage academic records"],["3","Review performance"],["4","Plan schedule and exams"],["5","Use suggestions and timer"]]; steps.forEach((a,i)=>{const x=72+i*224; box(s,x,282,180,130,i===2?navy:pale); text(s,a[0],x+24,306,45,35,24,i===2?yellow:blue,true); text(s,a[1],x+24,354,132,38,18,i===2?white:ink,true); if(i<4) text(s,"→",x+187,325,30,35,26,blue,true,"center");}); text(s,"This flow connects the source modules in a simple, diploma-project-friendly sequence.",72,505,1000,35,22,muted,false); }
// 13
{ const s=deck.slides.add(); top(s,"Future enhancements","Future scope",13); three(s,[["Mobile app & AI","Extend the platform to mobile access and explore AI-based performance prediction."],["Notifications & reminders","Add automated reminders for upcoming exams, study sessions and performance alerts."],["Parent panel & voice AI","Explore a parent access panel and voice-assistant support for easier navigation."]]); }
// 14
{ const s=deck.slides.add(); top(s,"Conclusion","Conclusion",14); text(s,"SAPPT is a focused academic-monitoring project that combines records, performance analysis and study planning in one web-based application.",72,190,1050,82,28,ink,false); box(s,72,326,1136,150,navy); text(s,"Core value",108,354,220,28,19,yellow,true); text(s,"Help students track progress, identify weak areas and organize their academic preparation.",108,395,960,40,25,white,true); text(s,"The project remains practical in scope while creating a clear foundation for future enhancements.",72,536,1030,34,22,muted,false); }
// 15
{ const s=deck.slides.add(); s.background.fill=navy; box(s,0,0,1280,18,yellow,"none"); text(s,"SAPPT V1 · Diploma Project",72,78,380,22,14,yellow,true); text(s,"Thank you",72,176,600,70,52,white,true); text(s,"Questions and discussion",72,265,550,36,25,"#DCE8FB",false); box(s,72,370,92,5,yellow,"none"); text(s,"Anuj Arun Deshmukh · Vishal Balande · Vishal Dake",72,450,800,32,22,white,true); }

const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(output);

