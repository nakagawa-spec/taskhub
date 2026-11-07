/* 実際の表や外部の通信は使わない。画面と保存処理の確認専用。 */
(function(g){
  const clone=x=>JSON.parse(JSON.stringify(x));
  class Sheet {
    constructor(rows){this.rows=clone(rows);}
    getLastRow(){return this.rows.length;}
    getDataRange(){return this.getRange(1,1,this.rows.length,Math.max(...this.rows.map(r=>r.length)));}
    getRange(row,col,n=1,m=1){const self=this;return {
      getValues(){return Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>(self.rows[row+i-1]||[])[col+j-1]??''));},
      setValues(values){values.forEach((r,i)=>r.forEach((v,j)=>{self.rows[row+i-1]||=[];self.rows[row+i-1][col+j-1]=v;}));return this;},
      getValue(){return this.getValues()[0][0];},setValue(v){return this.setValues([[v]]);},setFontWeight(){return this;},setNumberFormat(){return this;}
    };}
    appendRow(row){this.rows.push(clone(row));}
    setFrozenRows(){}
  }
  const members=[{email:'nakagawa@example.jp',name:'中川'},{email:'sato@example.jp',name:'佐藤'},{email:'suzuki@example.jp',name:'鈴木'}];
  const state={email:members[0].email,owner:members[0].email,lock:false,sequence:0,props:{},members};
  const sheets={tasks:new Sheet([['id','json','updatedAt','deleted','user']]),users:new Sheet([['email','name','registeredAt']])};
  const memberSheet=new Sheet([['email','name','role','active','joined_at'],...members.map(m=>[m.email,m.name,'member',true,''])]);
  function add(id,v,user){sheets.tasks.rows.push([id,JSON.stringify(v),++state.sequence,false,user]);}
  ['新サービスの発表','見学会の準備','社内報づくり'].forEach((name,i)=>{
    const id='kp_demo'+i, leader=members[i];
    add(id,{kind:'kouhouProject',name,client:i===0?'青空制作会社':'',leader:leader.email,leaderName:leader.name,notes:i===0?'広報室と制作会社で、公開までの準備を共有します。':'担当と期限を確認して進めましょう。',order:i,createdAt:'2026-11-06T00:00:00Z'},'__kouhou__');
    const count=i===0?4:2;
    for(let j=0;j<count;j++){
      const type=i===0&&j===2?'client':i===0&&j===3?'other':'member',m=members[j%3],assignee=type==='member'?m.name:type==='client'?'青空制作会社':'山田（撮影担当）';
      add('kt_demo'+i+'_'+j,{title:['告知内容をまとめる','写真と原稿を確認する','デザインの初稿を提出する','会場の撮影を手配する'][j],kouhouProjectId:id,kouhouProjectName:name,assigneeType:type,assigneeEmail:type==='member'?m.email:'',assigneeName:assignee,owner:type==='client'?'先方：'+assignee:assignee,startDate:'2026-11-06',dueDate:'2026-11-'+String(10+i+j),notes:j===0?'内容を確認して、打ち合わせで共有してください。':'',completed:j===1,recurrence:'none',priority:'none',order:j,createdAt:'2026-11-06T00:00:00Z'},type==='member'?m.email:leader.email);
    }
  });
  g.__kpTest={state,sheets,memberSheet,add,Sheet};
  g.Session={getActiveUser:()=>({getEmail:()=>state.email}),getEffectiveUser:()=>({getEmail:()=>state.owner})};
  g.SpreadsheetApp={getActiveSpreadsheet:()=>({getSheetByName:name=>sheets[name],insertSheet:name=>(sheets[name]=new Sheet([]))}),openById:()=>({getSheetByName:name=>name==='members'?memberSheet:null}),flush(){}};
  g.PropertiesService={getScriptProperties:()=>({getProperty:k=>state.props[k]??null,setProperty:(k,v)=>state.props[k]=v,deleteProperty:k=>delete state.props[k]})};
  g.LockService={getScriptLock:()=>({tryLock:()=>{if(state.lock)throw new Error('二重にロックしました');state.lock=true;return true;},releaseLock:()=>state.lock=false})};
  g.ScriptApp={getService:()=>({getUrl:()=> '#試しページ'})};
  g.Utilities={getUuid:()=> 'test-'+(++state.sequence),formatDate:()=> '2026-11-06'};
  g.Logger={log(){}};
  g.HtmlService={createHtmlOutputFromFile:name=>({name,setTitle(){return this;},addMetaTag(){return this;}}),createHtmlOutput:text=>({text,setTitle(){return this;}})};
})(globalThis);
