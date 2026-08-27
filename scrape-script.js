(function(){
function u8(s){return new TextEncoder().encode(s)}
function crc(b){var c,t=[],k,n;for(n=0;n<256;n++){c=n;for(k=0;k<8;k++)c=(c&1)?0xEDB88320^(c>>>1):(c>>>1);t[n]=c>>>0}var r=0xFFFFFFFF;for(var i=0;i<b.length;i++)r=t[(r^b[i])&255]^(r>>>8);return(r^0xFFFFFFFF)>>>0}
function zip(fs){var ch=[],ce=[],off=0,i;for(i=0;i<fs.length;i++){var nm=u8(fs[i].name),dt=fs[i].data,cr=crc(dt);var lh=new DataView(new ArrayBuffer(30));lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0x0800,true);lh.setUint32(14,cr,true);lh.setUint32(18,dt.length,true);lh.setUint32(22,dt.length,true);lh.setUint16(26,nm.length,true);ch.push(new Uint8Array(lh.buffer),nm,dt);var cd=new DataView(new ArrayBuffer(46));cd.setUint32(0,0x02014b50,true);cd.setUint16(4,20,true);cd.setUint16(6,20,true);cd.setUint16(8,0x0800,true);cd.setUint32(16,cr,true);cd.setUint32(20,dt.length,true);cd.setUint32(24,dt.length,true);cd.setUint16(28,nm.length,true);cd.setUint32(42,off,true);ce.push(new Uint8Array(cd.buffer),nm);off+=30+nm.length+dt.length}var co=off,cs=0;for(i=0;i<ce.length;i++)cs+=ce[i].length;var eo=new DataView(new ArrayBuffer(22));eo.setUint32(0,0x06054b50,true);eo.setUint16(8,fs.length,true);eo.setUint16(10,fs.length,true);eo.setUint32(12,cs,true);eo.setUint32(16,co,true);var all=ch.concat(ce,[new Uint8Array(eo.buffer)]),tot=0,p=0;for(i=0;i<all.length;i++)tot+=all[i].length;var out=new Uint8Array(tot);for(i=0;i<all.length;i++){out.set(all[i],p);p+=all[i].length}return out}
function xmlesc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}
function xmlRows(rows){var sd="<sheetData>";for(var i=0;i<rows.length;i++){sd+="<row r=\""+(i+1)+"\">";for(var j=0;j<rows[i].length;j++){var ref=String.fromCharCode(65+j)+(i+1);var v=rows[i][j];if(typeof v==="number"){sd+="<c r=\""+ref+"\" t=\"n\"><v>"+v+"</v></c>"}else{sd+="<c r=\""+ref+"\" t=\"inlineStr\"><is><t>"+xmlesc(v)+"</t></is></c>"}}sd+="</row>"}return sd+"</sheetData>"}
function xlsx(rows){var cw="<cols><col min=\"1\" max=\"1\" width=\"24\"/><col min=\"2\" max=\"2\" width=\"42\"/><col min=\"3\" max=\"3\" width=\"10\"/><col min=\"4\" max=\"4\" width=\"10\"/><col min=\"5\" max=\"5\" width=\"10\"/></cols>";var sd=xmlRows(rows);var ct="<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>";var rl="<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";var wb="<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"菜单\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";var wbr="<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>";return zip([{name:"[Content_Types].xml",data:u8(ct)},{name:"_rels/.rels",data:u8(rl)},{name:"xl/workbook.xml",data:u8(wb)},{name:"xl/_rels/workbook.xml.rels",data:u8(wbr)},{name:"xl/worksheets/sheet1.xml",data:u8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"+cw+sd+"</worksheet>")}])}
function esc(s){return String(s||"").replace(/[\n\r]/g," ").replace(/\s+/g," ").trim()}
function num(s){var m=String(s||"").match(/(\d+\.?\d*)/);return m?parseFloat(m[1]):0}
var groups=[],cur="";
var all=document.querySelectorAll("div,li,a,section");
for(var i=0;i<all.length;i++){
  var el=all[i],t=esc(el.innerText);
  if(!t||t.length>40)continue;
  var img=el.querySelector("img");
  var isHead=el.children.length<=2&&!img&&t.length<=10&&/^(全部|招牌|热销|推荐|人气|主食|套餐|小炒|米饭|炒饭|面|拌面|粉|粥|饮品|饮料|时蔬|素菜|荤菜|凉菜|小菜|小吃|加购|折扣|汤|甜点|新品|一人食)/.test(t);
  if(isHead&&!cur){cur=t;continue}
  if(cur&&img&&/￥|¥|\d+\.\d+/.test(t)){
    groups.push({g:cur,n:esc(t),p:num(t),img:img.src});
  }
}
var seen={},uniq=[];
groups.forEach(function(x){var k=x.g+"|"+x.n;if(!seen[k]){seen[k]=1;uniq.push(x)}});
var rows=[["分组","菜品名称","价格(元)","月售","有无图片"]];
uniq.forEach(function(x){
  var m=x.n.match(/月售\s*(\d+)/)||x.n.match(/月销\s*(\d+)/)||x.n.match(/已售\s*(\d+)/);
  rows.push([x.g,x.n,x.p,m?m[1]:"",x.img?"有":""]);
});
if(uniq.length<2){alert("未识别到菜单，请确认已进入店铺菜单页面后重试（脚本需在店铺菜单页执行）");return}
var blob=xlsx(rows);
var a=document.createElement("a");a.href=URL.createObjectURL(new Blob([blob],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));a.download="抓取菜单_"+Date.now()+".xlsx";document.body.appendChild(a);a.click();setTimeout(function(){document.body.removeChild(a)},300);
alert("抓取完成："+uniq.length+" 个菜品，已下载 Excel。若菜品不全，请在店铺页向下滚动加载全部菜单后重新运行本脚本。");
})();
