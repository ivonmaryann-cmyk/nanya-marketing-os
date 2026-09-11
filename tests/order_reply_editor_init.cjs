const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
function check(preloaded, selectionOnly=false) {
  const events = {}, frameEvents = {}, barEvents = {};
  let prepended = 0;
  const body = {innerHTML:'<table></table>',style:{},setAttribute(){},prepend(){prepended++;}};
  const doc = {querySelectorAll(){return [];},head:{append(){}},getElementById(){return null;},createElement(){return {querySelector(){return null;},contains(){return false;},style:{},setAttribute(){}};},body:preloaded?body:null,addEventListener(k,v){events[k]=v;},getSelection(){return null;}};
  const frame={contentDocument:doc,addEventListener(k,v){frameEvents[k]=v;}};
  const bar={addEventListener(k,v){barEvents[k]=v;}};
  let alerts=[];
  vm.runInNewContext(fs.readFileSync('static/order_reply_editor.js','utf8'),{
    document:{getElementById:id=>id==='replyEditor'?frame:bar},alert:m=>alerts.push(m),
    prompt:()=>null,confirm:()=>false,setInterval(){}
  });
  if(!preloaded){doc.body=body;frameEvents.load();}
  assert.equal(body.contentEditable,'true');
  assert.equal(typeof events.click,'function');
  const table={rows:[]};
  const cell={isConnected:true,closest:()=>table};
  const nested={nodeType:1,closest:()=>cell};
  if(selectionOnly) doc.getSelection=()=>({anchorNode:{nodeType:3,parentElement:nested}});
  else events.click({target:nested});
  // delete-col does not require a position; cancellation confirms selection survived.
  barEvents.click({target:{closest:()=>({dataset:{action:'delete-col'}})}});
  assert.deepEqual(alerts,[]);
  frameEvents.load(); // Same document must not reset the selected cell.
  assert.equal(prepended,1);
  barEvents.click({target:{closest:()=>({dataset:{action:'delete-col'}})}});
  assert.deepEqual(alerts,[]);
}
check(true);check(false);check(true,true);check(false,true);
console.log('Editor early/late initialization and selected-cell retention passed');
