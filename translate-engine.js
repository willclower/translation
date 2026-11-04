/* translate-engine.js — v2 (2026-11-04)
 * One job: turn a .docx into whole-paragraph translation units, put the translations back
 * with formatting and links intact, and check every paragraph in code before delivery.
 *
 *  - extractUnits(zip)          -> { units[], parts{} }   one unit per paragraph with text
 *  - checkUnit(unit, text, lang)-> { issues[], parsed }   code checks, no model
 *  - rebuildDoc(zip, parts, units, results) -> Blob/Buffer  (results[i] = {text, issues[]})
 *
 * A paragraph with several formatting runs or a link is sent as
 *   "⟦1⟧You can get the ⟦/1⟧⟦2⟧recommended⟦/2⟧⟦3⟧ amount of activity.⟦/3⟧"
 * so the model translates the whole sentence and code puts each part back on its run.
 * Browser and Node (tests) both load this file.
 */
(function (root) {
  'use strict';
  var ENGINE_VERSION = 'engine-v3';
  var DOC_RE = /^word\/(document|header\d*|footer\d*)\.xml$/;

  function xmlDecode(s){return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');}
  function xmlEncode(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
  function normRpr(r){ // formatting identity of a run; ignore proofing-only differences
    return (r||'').replace(/<w:lang\b[^>]*\/>/g,'').replace(/<w:noProof\b[^>]*\/>/g,'')
      .replace(/<w:spellcheck\b[^>]*\/>/g,'').replace(/\sw:hint="[^"]*"/g,'').replace(/\s+/g,' ').trim();
  }
  var MK = /⟦(\/?)(\d+)⟧/g;
  function stripMarkers(s){ return String(s==null?'':s).replace(MK,''); }

  // ---------- extract ----------
  function scanPart(xml){
    var TAG=/<(\/?)(w:p|w:hyperlink|w:r|w:rPr|w:t)(?=[\s>\/])([^>]*?)(\/?)>/g, m;
    var pStack=[], hStack=[], rStack=[], paras=[], slots=[], pid=0, hid=0;
    while((m=TAG.exec(xml))!==null){
      var close=m[1]==='/', name=m[2], selfClose=m[4]==='/';
      if(name==='w:p'){
        if(selfClose) continue;
        if(!close){ var p={id:pid++, openEnd:m.index+m[0].length, closeAt:-1, depth:pStack.length, slots:[]}; paras.push(p); pStack.push(p); }
        else { var pp=pStack.pop(); if(pp) pp.closeAt=m.index; }
      } else if(name==='w:hyperlink'){
        if(selfClose) continue;
        if(!close) hStack.push(hid++); else hStack.pop();
      } else if(name==='w:r'){
        if(selfClose) continue;
        if(!close) rStack.push({rpr:null, rprStart:-1}); else rStack.pop();
      } else if(name==='w:rPr'){
        var run=rStack[rStack.length-1];
        if(!run || run.rpr!==null) continue;          // paragraph-mark rPr or already set
        if(selfClose){ run.rpr=''; continue; }
        if(!close) run.rprStart=m.index+m[0].length; else if(run.rprStart>=0) run.rpr=xml.slice(run.rprStart,m.index);
      } else if(name==='w:t'){
        if(close||selfClose) continue;
        var cs=m.index+m[0].length, ce=xml.indexOf('</w:t>',cs);
        if(ce<0) break;
        var para=pStack[pStack.length-1], r=rStack[rStack.length-1];
        if(para){
          var sl={tagStart:m.index, contentStart:cs, contentEnd:ce, text:xmlDecode(xml.slice(cs,ce)),
                  key:normRpr(r&&r.rpr)+'|h'+(hStack.length?hStack[hStack.length-1]:'')};
          para.slots.push(sl); slots.push(sl);
        }
        TAG.lastIndex=ce+6;
      }
    }
    return paras;
  }

  async function extractUnits(zip){
    var units=[], parts={};
    var names=Object.keys(zip.files).filter(function(n){return DOC_RE.test(n);}).sort();
    for(var k=0;k<names.length;k++){
      var xml=await zip.file(names[k]).async('string');
      var paras=scanPart(xml);
      parts[names[k]]={xml:xml, paras:paras};
      for(var j=0;j<paras.length;j++){
        var p=paras[j]; if(!p.slots.length) continue;
        var spans=[], cur=null;
        p.slots.forEach(function(s){
          if(cur && cur.key===s.key){ cur.slots.push(s); cur.text+=s.text; }
          else { cur={key:s.key, slots:[s], text:s.text}; spans.push(cur); }
        });
        var live=spans.filter(function(s){return s.text.trim()!=='';});
        if(!live.length) continue;
        live.forEach(function(s,i){ s.n=i+1; });
        var send = live.length===1 ? live[0].text.trim()
                 : live.map(function(s){ return '⟦'+s.n+'⟧'+s.text+'⟦/'+s.n+'⟧'; }).join('');
        units.push({part:names[k], para:p, spans:live, marked:live.length>1, send:send,
                    plain:live.map(function(s){return s.text;}).join('')});
      }
    }
    return {units:units, parts:parts};
  }

  // ---------- parse a translated unit back into spans ----------
  function parseUnit(unit, text){
    text=String(text==null?'':text);
    if(!unit.marked) return {ok:true, moved:false, texts:[stripMarkers(text)]};
    var n=unit.spans.length, texts=new Array(n+1), order=[], open=null, last=null, pending='', ok=true, idx=0, m;
    MK.lastIndex=0;
    function addLoose(t){ if(!t) return; if(open!==null) texts[open]+=t; else if(last!==null) texts[last]+=t; else pending+=t; }
    while((m=MK.exec(text))!==null){
      addLoose(text.slice(idx,m.index)); idx=m.index+m[0].length;
      var k=+m[2];
      if(k<1||k>n){ ok=false; continue; }
      if(m[1]!=='/'){ if(open!==null||texts[k]!==undefined){ ok=false; continue; } open=k; texts[k]=pending; pending=''; order.push(k); }
      else { if(open!==k){ ok=false; continue; } last=k; open=null; }
    }
    addLoose(text.slice(idx));
    if(open!==null) ok=false;
    for(var i=1;i<=n;i++) if(texts[i]===undefined) ok=false;
    if(!ok) return {ok:false, moved:false, texts:[stripMarkers(text)]};
    var moved=order.some(function(k,i){return k!==i+1;});
    // text order wins over formatting: returned order goes onto the runs in document order
    return {ok:true, moved:moved, texts:order.map(function(k){return texts[k];})};
  }

  // ---------- code checks (no model) ----------
  var EN_STOP=['the','and','of','to','with','your','you','for','is','are','this','that','it','from','have','will',
    'be','by','or','these','their','they','which','should','when','what','how','our','we','can\'t','don\'t'];
  var METRIC_T=/(\bkg\b|\bkilo(s|gramo|gramos)?\b|\bgramos?\b|\bkm\b|\bkil[oó]metros?\b|\bcm\b|\bcent[ií]metros?\b|\bmetros?\b|\blitros?\b|°\s?C\b|\bcelsius\b)/i;
  var METRIC_S=/(\bkg\b|kilogram|\bgrams?\b|\bkm\b|kilomet|centimet|\bcm\b|\bmeters?\b|\bmetres?\b|\blit(er|re)s?\b|°\s?C\b|celsius)/i;
  var US_UNITS_ES=[ // [source pattern, Spanish pattern, label]
    [/\b(pounds?|lbs?\.?)\b/i, /\b(libras?|lbs?\.?)\b/i, 'pounds (libras)'],
    [/\b(ounces?|oz\.?)\b/i, /\b(onzas?|oz\.?)\b/i, 'ounces (onzas)'],
    [/\bmiles?\b/i, /\bmillas?\b/i, 'miles (millas)'],
    [/\b(feet|foot|ft\.?)\b/i, /\b(pies?|ft\.?)\b/i, 'feet (pies)'],
    [/\b(inch|inches)\b/i, /\bpulgadas?\b/i, 'inches (pulgadas)'],
    [/°\s?F\b|\bfahrenheit\b/i, /°\s?F\b|\bfahrenheit\b/i, '°F']
  ];
  var ROMANCE=/^(spanish|french|portuguese|italian)/i;

  function nums(s){ return (s.match(/\d+(?:[.,\s]\d{3})*(?:[.,]\d+)?/g)||[]).map(function(x){return x.replace(/[.,\s]/g,'');}); }
  function countChar(s,c){ return s.split(c).length-1; }
  function noUrls(s){ return s.replace(/https?:\/\/\S+|www\.\S+|\S+@\S+\.\S+/g,' '); }

  function checkUnit(unit, text, lang){
    var issues=[], parsed=parseUnit(unit, text);
    var src=unit.plain, tgt=parsed.texts.join('');
    var srcT=noUrls(src), tgtT=noUrls(tgt);
    if(!parsed.ok) issues.push('formatting/link markers came back broken — formatting in this paragraph may be off');
    else if(parsed.moved) issues.push('formatting or link may sit on the wrong words (the sentence order changed)');
    if(parsed.ok && unit.marked){
      parsed.texts.forEach(function(t,i){ // link text much shorter or longer than its English = likely on the wrong words
        var sp=unit.spans[i], st=sp.text.trim(), tt=t.trim();
        if(!/\|h$/.test(sp.key) && tt && st.length>=12 && (tt.length/st.length<0.55 || tt.length/st.length>2.2))
          issues.push('the link may be on the wrong words ("'+tt.slice(0,40)+'" for "'+st.slice(0,40)+'")'); });
      parsed.texts.forEach(function(t,i){ if(!t.trim()) issues.push('the '+(/\|h$/.test(unit.spans[i].key)?'formatted':'link')+' text "'+unit.spans[i].text.trim().slice(0,40)+'" came back empty'); });
    }
    var words=(srcT.match(/[A-Za-z]+/g)||[]);
    var stopHit=function(s){ return ((s.match(/[A-Za-zÀ-ÿ']+/g)||[]).filter(function(w){return EN_STOP.indexOf(w)>=0;})); }; // lowercase only: names like "Walk With Ease" pass
    if(tgt.trim()===src.trim() && words.length>=3 && stopHit(src).length) issues.push('not translated (same as the English)');
    else { var left=stopHit(tgtT); if(left.length) issues.push('English words left in: '+left.slice(0,5).join(', ')); }
    var a=nums(srcT), b=nums(tgtT), bb=b.slice(), aa=a.slice();
    a.forEach(function(x){ var i=bb.indexOf(x); if(i>=0) bb.splice(i,1); else issues.push('number '+x+' missing'); });
    b.forEach(function(x){ var i=aa.indexOf(x); if(i>=0) aa.splice(i,1); });
    if(bb.length) issues.push('new number(s) not in the English: '+bb.slice(0,4).join(', '));
    ['$','%'].forEach(function(c){ if(countChar(srcT,c)!==countChar(tgtT,c)) issues.push('"'+c+'" count differs from the English'); });
    if(METRIC_T.test(tgtT) && !METRIC_S.test(srcT)) issues.push('converted to metric — keep US units');
    if(/^spanish/i.test(lang||'') && /\b(pounds?|ounces?|miles|inch(es)?|feet)\b/i.test(tgtT)) issues.push('English unit word left in (write libras, onzas, millas, pulgadas, pies)');
    if(/^spanish/i.test(lang||'')) US_UNITS_ES.forEach(function(u){ if(u[0].test(srcT) && !u[1].test(tgtT)) issues.push(u[2]+' missing'); });
    var sl=src.trim().length, tl=tgt.trim().length;
    if(ROMANCE.test(lang||'') && sl>=40 && tl/sl<0.8) issues.push('much shorter than the English — something may be missing');
    if(sl>=20 && tl/sl>2.2) issues.push('much longer than the English — something may have been added');
    var fused=(tgtT.match(/[A-Za-zÀ-ÿ]{25,}/g)||[]).filter(function(w){return srcT.indexOf(w)<0;});
    if(fused.length) issues.push('possible words run together: '+fused[0]);
    return {issues:issues, parsed:parsed};
  }

  // ---------- rebuild ----------
  function edges(orig, core){
    var lead=(orig.match(/^\s*/)||[''])[0], trail=(orig.match(/\s*$/)||[''])[0];
    return lead+String(core).trim()+trail;
  }
  function rebuildXml(part, units, results, flagFor){
    var xml=part.xml, edits=[];
    units.forEach(function(u, ui){
      var r=results[ui]; if(!r) return;
      var parsed=r.parsed, texts=parsed.texts;
      u.spans.forEach(function(sp, si){
        var t = parsed.ok ? (texts[si]||'') : (si===0 ? texts[0] : '');
        t = (t.trim()==='' && !parsed.ok && si>0) ? '' : edges(sp.text, t);
        sp.slots.forEach(function(sl, k){
          edits.push({at:sl.tagStart, del:sl.contentEnd-sl.tagStart,
                      ins:'<w:t xml:space="preserve">'+(k===0?xmlEncode(t):'')});
        });
      });
      var flags=(r.issues||[]).concat(r.notes||[]);
      if(flagFor && flags.length && u.para.depth===0 && u.para.closeAt>0){
        var id=flagFor(flags), at=u.para.openEnd;
        if(xml.startsWith('<w:pPr',at)){ var e=xml.indexOf('</w:pPr>',at), sc=xml.indexOf('/>',at), gt=xml.indexOf('>',at);
          at = (sc>=0 && sc+1===gt) ? gt+1 : e+8; }
        edits.push({at:at, del:0, ins:'<w:commentRangeStart w:id="'+id+'"/>'});
        edits.push({at:u.para.closeAt, del:0, ins:'<w:commentRangeEnd w:id="'+id+'"/><w:r><w:commentReference w:id="'+id+'"/></w:r>'});
        r.commented=true;
      }
    });
    edits.sort(function(a,b){ return b.at-a.at || a.del-b.del; });
    edits.forEach(function(e){ xml=xml.slice(0,e.at)+e.ins+xml.slice(e.at+e.del); });
    return xml;
  }

  async function rebuildDoc(zip, parts, units, results, opts){
    opts=opts||{};
    var comments=[], nextId=null;
    if(opts.comments){
      var maxId=0, docXml=parts['word/document.xml']?parts['word/document.xml'].xml:'';
      var cf=zip.file('word/comments.xml'), cxml=cf?await cf.async('string'):null;
      (docXml+(cxml||'')).replace(/w:id="(\d+)"/g,function(_,d){ maxId=Math.max(maxId,+d); });
      nextId=maxId+1;
    }
    var names=Object.keys(parts);
    for(var k=0;k<names.length;k++){
      var name=names[k], us=[], rs=[];
      units.forEach(function(u,i){ if(u.part===name){ us.push(u); rs.push(results[i]); } });
      var flagFor = (opts.comments && name==='word/document.xml') ? function(issues){
        var id=nextId++; comments.push({id:id, text:'Translation check: '+issues.join('; ').replace(/\.$/,'')+'.'}); return id; } : null;
      zip.file(name, rebuildXml(parts[name], us, rs, flagFor));
    }
    if(comments.length) await addComments(zip, comments);
    return zip.generateAsync({type: opts.type||'blob', compression:'DEFLATE'});
  }

  async function addComments(zip, list){
    var date=new Date().toISOString().replace(/\.\d+Z$/,'Z');
    var body=list.map(function(c){ return '<w:comment w:id="'+c.id+'" w:author="MedWell translation check" w:initials="MW" w:date="'+date+'"><w:p><w:r><w:t xml:space="preserve">'+xmlEncode(c.text)+'</w:t></w:r></w:p></w:comment>'; }).join('');
    var cf=zip.file('word/comments.xml');
    if(cf){ var x=await cf.async('string'); zip.file('word/comments.xml', x.replace('</w:comments>', body+'</w:comments>')); return; }
    zip.file('word/comments.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+body+'</w:comments>');
    var rp='word/_rels/document.xml.rels', rels=await zip.file(rp).async('string');
    if(rels.indexOf('relationships/comments"')<0)
      zip.file(rp, rels.replace('</Relationships>','<Relationship Id="rIdMwCheck1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
    var ct=await zip.file('[Content_Types].xml').async('string');
    if(ct.indexOf('/word/comments.xml')<0)
      zip.file('[Content_Types].xml', ct.replace('</Types>','<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  }

  // ---------- translate with checks and one retry ----------
  // post(body) -> {translated[]} ; calls translate-doc
  async function translateChecked(units, lang, register, post, onStatus){
    var send=units.map(function(u){return u.send;});
    var tr=await post({mode:'paragraph', language:lang, register:register, strings:send});
    var out=(tr&&tr.translated)||tr;
    if(!Array.isArray(out)||out.length!==send.length) throw new Error('translation mismatch');
    var results=out.map(function(t,i){ var c=checkUnit(units[i], t, lang); return {text:t, issues:c.issues, parsed:c.parsed}; });
    var redo=[]; results.forEach(function(r,i){ if(r.issues.length) redo.push(i); });
    if(redo.length){
      if(onStatus) onStatus('rechecking '+redo.length+'\u2026');
      var tr2=await post({mode:'paragraph', language:lang, register:register,
        strings:redo.map(function(i){return send[i];}),
        notes:redo.map(function(i){return 'A previous attempt had these problems: '+results[i].issues.join('; ')+'. Fix them.';})});
      var out2=(tr2&&tr2.translated)||tr2;
      if(Array.isArray(out2) && out2.length===redo.length){
        redo.forEach(function(i,j){
          var c=checkUnit(units[i], out2[j], lang);
          if(c.issues.length < results[i].issues.length) results[i]={text:out2[j], issues:c.issues, parsed:c.parsed, retried:true};
          else results[i].retried=true;
        });
      }
    }
    var flagged=results.filter(function(r){return r.issues.length;}).length;
    return {results:results, flagged:flagged, retried:redo.length};
  }

  // ---------- second opinion (translate-review) ----------
  // A second model returns only paragraphs needing a change. Code re-checks every suggestion:
  //  grammar/style fixes that pass the checks are applied quietly (listed on the status line);
  //  meaning fixes are applied AND get a comment so a person confirms;
  //  a suggestion that fails the checks is not applied and gets a comment.
  async function reviewChecked(units, run, lang, register, post){
    var items=units.map(function(u,i){ return {i:i, source:u.send, translation:run.results[i].text}; });
    var res; try{ res=await post({language:lang, register:register, items:items}); }catch(e){ return {ok:false, applied:0, error:e.message}; }
    var list=(res&&res.reviews)||[], applied=0, changes=[];
    list.forEach(function(rv){
      var i=rv.i, r=run.results[i]; if(!r || typeof rv.corrected!=='string') return;
      if(rv.corrected.trim()===String(r.text).trim()) return;
      var c=checkUnit(units[i], rv.corrected, lang);
      var why=(rv.kind||'style')+': '+(rv.problem||'');
      if(c.issues.length<=r.issues.length){
        r.text=rv.corrected; r.parsed=c.parsed; r.issues=c.issues; r.reviewed=why; applied++;
        changes.push(i);
        if(rv.kind==='meaning'){ r.notes=(r.notes||[]).concat(['review changed the meaning-related wording, please confirm ('+rv.problem+')']); }
      } else {
        r.notes=(r.notes||[]).concat(['review suggests a change that failed the checks, not applied ('+rv.problem+') — suggested: '+stripMarkers(rv.corrected).slice(0,200)]);
      }
    });
    run.flagged=run.results.filter(function(r){ return r.issues.length || (r.notes&&r.notes.length); }).length;
    return {ok:!(res&&res.failed_batches&&res.failed_batches.length), applied:applied, changes:changes};
  }

  var api={ENGINE_VERSION:ENGINE_VERSION, extractUnits:extractUnits, parseUnit:parseUnit, checkUnit:checkUnit,
           rebuildDoc:rebuildDoc, translateChecked:translateChecked, reviewChecked:reviewChecked, stripMarkers:stripMarkers};
  if(typeof module!=='undefined' && module.exports) module.exports=api; else root.TranslateEngine=api;
})(this);
