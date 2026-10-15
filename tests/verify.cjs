const fs=require('fs'),vm=require('vm'),assert=require('assert'),path=require('path');
const root=path.resolve(__dirname,'..'),appRoot=fs.existsSync(path.join(root,'dist'))?path.join(root,'dist'):root,els=new Map();
function el(key){if(!els.has(key))els.set(key,{value:'',checked:true,style:{},dataset:{},classList:{toggle(){}},closest(){return this}});return els.get(key)}
let RealPptx;try{RealPptx=require('pptxgenjs')}catch{RealPptx=require('/opt/codex/runtimes/codex-primary-runtime/dependencies/node/node_modules/pptxgenjs/dist/pptxgen.cjs.js')}
let pptBlob;class Pptx extends RealPptx{async writeFile(){pptBlob=await this.write({outputType:'nodebuffer'})}}
const context={document:{querySelector:el,querySelectorAll:()=>[],addEventListener(){}},navigator:{},location:{protocol:"https:"},window:{scrollTo(){}},PptxGenJS:Pptx,console,setTimeout:()=>0};
vm.createContext(context);vm.runInContext(fs.readFileSync(path.join(appRoot,'app.js'),'utf8'),context);
const good=JSON.parse(fs.readFileSync(path.join(appRoot,'sample.json'),'utf8'));
context.input=structuredClone(good);vm.runInContext('lesson=validate(input);sample=true;render()',context);
assert.equal(context.qualityIssues(good).length,0);assert.equal(good.slides.length,20);
assert.match(el('#preview').innerHTML,/ここはプリントに記入/);assert.match(el('#preview').innerHTML,/考える時間 90秒/);
const student=vm.runInContext('printBody(false)',context),teacher=vm.runInContext('printBody(true)',context);
assert.match(student,/対応スライド/);assert(!student.includes('解答：'));assert(teacher.includes('解答：'));assert.match(teacher,/ここはプリントに記入/);
for(const mutate of [o=>o.slides[0].question=999,o=>o.slides[0].minutes=20,o=>o.slides.forEach(s=>{if(s.question===1)s.write=false}),o=>o.slides[8].thinkSeconds=900,o=>o.slides[4].example='']){const bad=structuredClone(good);mutate(bad);assert.throws(()=>context.validate(bad))}
const old=structuredClone(good);old.slides.forEach(s=>{delete s.purpose;delete s.write;delete s.activity;delete s.thinkSeconds;delete s.example});context.validate(old);
const short=structuredClone(good);short.slides=short.slides.slice(0,17);assert(context.qualityIssues(short).some(x=>x.includes('17枚')));
const page={text:'手動修正済みの文章',parts:[],reviewed:true};
context.applyOCR(page,{text:'',warnings:['空欄'],confidence:0,failed:true,parts:[]});assert.equal(page.text,'手動修正済みの文章');assert.equal(page.reviewed,false);assert.equal(page.ocrState,'失敗');
context.failedOCR(page,Error('通信失敗'));assert(page.warnings.some(x=>x.includes('通信失敗')));assert.equal(page.text,'手動修正済みの文章');
context.applyOCR(page,{text:'健康と生活習慣の関係を学ぶ',warnings:[],confidence:80,failed:false,parts:[{type:'body',text:'健康と生活習慣の関係を学ぶ'}]});assert(page.text.includes('【本文】'));assert.equal(page.reviewed,false);assert.equal(page.ocrState,'読み取り候補（未確認）');
vm.runInContext("photos=[{text:'未確認の日本語文章',reviewed:false}];sample=false;",context);el('#confirmed').checked=true;el('#unknown').value='';context.step(2);assert.equal(el('#toast').textContent,'全ページの本文と不明点を確認してください');vm.runInContext('sample=true;',context);
assert(context.makePrompt().includes('標準20枚'));assert(context.makePrompt().includes('引き延ば'));
(async()=>{await el('#pptx').onclick();assert(pptBlob&&pptBlob.length>1000,'PowerPoint export');fs.writeFileSync('/tmp/lesson-upgrade-check.pptx',pptBlob);console.log('PASS: 20 slides, 50 min, worksheet mappings, legacy JSON, content checks and actual PPTX export')})().catch(e=>{console.error(e);process.exitCode=1});
