/* 餐谋长·运营助手 - 商家调研零插件抓取脚本 v2（抓全版）
   由「抓取书签」在店铺页动态加载执行：
   1) 依次点击左侧全部分类tab，触发各分组菜品渲染
   2) 滚动到底加载全部菜品
   3) 提取 分组/名称/价格/券后价/月售/描述 -> 下载CSV(Excel可打开)
   仅抓取用户端公开展示的店铺菜单数据，用于商家调研分析。 */
(function(){
  'use strict';
  var BRAND='餐谋长·抓取';
  function banner(msg, color){
    var id='cmz-banner';
    var b=document.getElementById(id);
    if(!b){
      b=document.createElement('div');
      b.id=id;
      b.style.cssText='position:fixed;top:0;left:0;right:0;z-index:2147483647;padding:10px 16px;font-size:14px;font-weight:600;font-family:-apple-system,"PingFang SC",sans-serif;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.15);';
      document.documentElement.appendChild(b);
    }
    b.style.background=color||'#EFF6FF';
    b.style.color='#1D4ED8';
    b.style.borderBottom='2px solid #2563EB';
    b.innerHTML=msg;
  }
  function sleep(ms){ return new Promise(function(r){ setTimeout(r,ms); }); }
  function numText(s){ return String(s||'').replace(/[^\d.]/g,''); }
  function txt(item,sel){
    var el=item.querySelector(sel);
    return el?(el.innerText||'').replace(/\s+/g,'').trim():'';
  }
  function escCSV(s){
    s=String(s==null?'':s).replace(/\r?\n/g,' ');
    if(/[",\n]/.test(s)) s='"'+s.replace(/"/g,'""')+'"';
    return s;
  }
  function ts(){ var d=new Date(); function p2(x){return ('0'+x).slice(-2);} return d.getFullYear()+p2(d.getMonth()+1)+p2(d.getDate())+'_'+p2(d.getHours())+p2(d.getMinutes()); }
  function downloadCSV(text, name){
    var blob=new Blob(['\ufeff'+text],{type:'text/csv;charset=utf-8'});
    var url=URL.createObjectURL(blob);
    var a=document.createElement('a');
    a.href=url; a.download=name;
    document.body.appendChild(a); a.click();
    setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(url); },400);
  }
  function isPunish(){ return /punish|waimai-guide|x5secdata/i.test(location.href); }
  function parseMenu(){
    var items=[], seen={};
    document.querySelectorAll('.menuItem').forEach(function(item){
      var cate=item.getAttribute('data-cate-name')||'';
      var title=txt(item,'.menuItem--info-title')||txt(item,'.menuItem--info-title--warp');
      var price=txt(item,'.menuItem--info-price-text')||numText(txt(item,'.menuItem--info-price'));
      var use=txt(item,'.menuItem--info-useCouponPrice');
      var sales=txt(item,'.menuItem--info-sales');
      var desc=txt(item,'.menuItem--info-description');
      if(!title) return;
      var key=title+'|'+cate;
      if(seen[key]) return; seen[key]=1;
      items.push({cate:cate,name:title,price:price,use:use,sales:sales,desc:desc});
    });
    return items;
  }
  /* 依次点击左侧全部分类tab，触发各分组菜品渲染 */
  async function clickAllTabs(){
    var tabs=document.querySelectorAll('.sideList--item');
    if(!tabs.length) return true;
    for(var i=0;i<tabs.length;i++){
      if(isPunish()) return false;
      try{ tabs[i].click(); }catch(e){}
      await sleep(900+Math.random()*500);
    }
    return true;
  }
  /* 滚动到最底部，直到菜品数量稳定（全部加载完成） */
  async function loadAll(){
    var sc=document.querySelector('.mor-comp-page-content');
    if(!sc){
      var all=document.querySelectorAll('*');
      for(var i=0;i<all.length;i++){
        var el=all[i];
        if(el.scrollHeight>el.clientHeight+200 && el.clientHeight>100){ sc=el; break; }
      }
    }
    if(!sc) return true;
    var prev=-1, stable=0;
    for(var i=0;i<50 && stable<4; i++){
      sc.scrollTop=sc.scrollHeight;
      await sleep(700+Math.random()*500);
      if(isPunish()) return false;
      var n=document.querySelectorAll('.menuItem').length;
      if(n===prev){ stable++; } else { stable=0; prev=n; }
    }
    sc.scrollTop=0;
    return true;
  }
  (async function(){
    if(isPunish()){
      banner('⚠️ '+BRAND+'：当前是平台安全验证页。请先用手机「饿了么」App 扫码登录后，再重新打开店铺点书签。','#FEF3C7');
      return;
    }
    banner('⏳ '+BRAND+'：正在展开全部菜单（遍历分类+滚动加载），请稍候…');
    var waited=0;
    while(waited<30000){
      if(document.querySelectorAll('.food-item--wrap,.menuItem,.sideList--item').length>0) break;
      if(isPunish()){
        banner('⚠️ '+BRAND+'：触发平台安全验证，请登录后重试。','#FEF3C7');
        return;
      }
      await sleep(500); waited+=500;
    }
    var ok=await clickAllTabs();
    if(!ok){ banner('⚠️ '+BRAND+'：加载中触发安全验证，已停止。','#FEF3C7'); return; }
    ok=await loadAll();
    if(!ok){ banner('⚠️ '+BRAND+'：加载中触发安全验证，已停止。','#FEF3C7'); return; }
    if(isPunish()){
      banner('⚠️ '+BRAND+'：抓取中触发平台安全验证，已自动停止，请登录后重试。','#FEF3C7');
      return;
    }
    var items=parseMenu();
    if(!items.length){
      banner('❌ '+BRAND+'：未识别到菜单。请确认当前页面是店铺「点餐」页后，再点一次书签。','#FEF2F2');
      return;
    }
    var lines=['分组,菜品名称,现价(元),券后/预估到手(元),月售,说明'];
    items.forEach(function(x){
      lines.push([escCSV(x.cate),escCSV(x.name),escCSV(x.price),escCSV(x.use),escCSV(x.sales),escCSV(x.desc)].join(','));
    });
    var shop='';
    var t=document.body.innerText||'';
    var m=t.match(/([^\n]{2,20}?\([^)]*店\))/)||t.match(/[^\n]{2,15}店/);
    if(m) shop=String(m[1]||'').replace(/[\\\/:*?"<>|\n\r]/g,' ').trim().slice(0,30);
    downloadCSV(lines.join('\n'),'抓取菜单_'+(shop||'店铺')+'_'+ts()+'.csv');
    banner('✅ '+BRAND+'：识别到 '+items.length+' 个菜品（已遍历全部 '+items.filter(function(x){return x.cate;}).length+' 个分组），CSV 已自动下载，可用 Excel 打开。可关闭本页。','#F0FDF4');
  })();
})();