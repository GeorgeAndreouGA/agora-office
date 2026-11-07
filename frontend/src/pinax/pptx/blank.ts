/**
 * The built-in template for new presentations: a complete, minimal PresentationML package
 * (16:9) with a "Greek Office" theme, one master, six layouts, a notes master and one title
 * slide. Generated as text so it is easy to read and to keep valid.
 */
import { strToU8, zipSync } from 'fflate';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const W = 12192000;
const H = 6858000;

function rels(list: [string, string, string][]): string {
  return (
    HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('') +
    '</Relationships>'
  );
}

const THEME = (name: string) => `${HEAD}<a:theme ${NS} name="${name}"><a:themeElements>
<a:clrScheme name="Greek Office">
<a:dk1><a:srgbClr val="1B1712"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="2A231D"/></a:dk2><a:lt2><a:srgbClr val="F4EFE6"/></a:lt2>
<a:accent1><a:srgbClr val="A4472A"/></a:accent1><a:accent2><a:srgbClr val="4E5B23"/></a:accent2>
<a:accent3><a:srgbClr val="2E5A7C"/></a:accent3><a:accent4><a:srgbClr val="A8802F"/></a:accent4>
<a:accent5><a:srgbClr val="7E3420"/></a:accent5><a:accent6><a:srgbClr val="5C5248"/></a:accent6>
<a:hlink><a:srgbClr val="2E5A7C"/></a:hlink><a:folHlink><a:srgbClr val="7E3420"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Greek Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Greek Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="12700" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="19050" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="25400" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

function lvl(n: number, sz: number, marL: number, bullet: boolean, extra = ''): string {
  const bu = bullet
    ? `<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="${n % 2 ? '•' : '–'}"/>`
    : '<a:buNone/>';
  return `<a:lvl${n}pPr marL="${marL}" indent="${bullet ? -228600 : 0}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="${n === 1 ? 1000 : 500}"/></a:spcBef>${bu}<a:defRPr sz="${sz}" kern="1200"${extra}><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n}pPr>`;
}

const plainLevels = Array.from({ length: 9 }, (_, i) => {
  const n = i + 1;
  return `<a:lvl${n}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n}pPr>`;
}).join('');

const PRESENTATION = `${HEAD}<p:presentation ${NS} saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:notesMasterIdLst><p:notesMasterId r:id="rId3"/></p:notesMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst>
<p:sldSz cx="${W}" cy="${H}"/><p:notesSz cx="6858000" cy="9144000"/>
<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${plainLevels}</p:defaultTextStyle>
</p:presentation>`;

function ph(
  id: number,
  name: string,
  type: string | null,
  idx: number | null,
  x: number,
  y: number,
  cx: number,
  cy: number,
  body = '<a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>',
  extraNv = '',
): string {
  const attrs = [type ? `type="${type}"` : '', idx !== null ? `idx="${idx}"` : '', extraNv].filter(Boolean).join(' ');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm></p:spPr><p:txBody>${body}</p:txBody></p:sp>`;
}

function prompt(text: string, bodyPr = '<a:bodyPr/>', lstStyle = '<a:lstStyle/>'): string {
  return `${bodyPr}${lstStyle}<a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p>`;
}

const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

const FOOTERS =
  ph(5, 'Date Placeholder 4', 'dt', 10, 838200, 6356350, 2743200, 365125, prompt('', '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', '<a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>')) +
  ph(6, 'Footer Placeholder 5', 'ftr', 11, 4038600, 6356350, 4114800, 365125, prompt('', '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>')) +
  ph(7, 'Slide Number Placeholder 6', 'sldNum', 12, 8610600, 6356350, 2743200, 365125, `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`);

// A thin gold rule above the footer, the theme's one decoration.
const RULE = `<p:sp><p:nvSpPr><p:cNvPr id="8" name="Gold rule"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="6281928"/><a:ext cx="10515600" cy="19050"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent4"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;

const MASTER = `${HEAD}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}
${ph(2, 'Title Placeholder 1', 'title', null, 838200, 365125, 10515600, 1325563, prompt('Click to edit Master title style', '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>'))}
${ph(3, 'Text Placeholder 2', 'body', 1, 838200, 1825625, 10515600, 4351338, `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US"/><a:t>Third level</a:t></a:r></a:p>`)}
${FOOTERS}${RULE}
</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst>${[1, 2, 3, 4, 5, 6].map((i) => `<p:sldLayoutId id="${2147483648 + i}" r:id="rId${i}"/>`).join('')}</p:sldLayoutIdLst>
<p:txStyles>
<p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>
<p:bodyStyle>${lvl(1, 2800, 228600, true)}${lvl(2, 2400, 685800, true)}${lvl(3, 2000, 1143000, true)}${lvl(4, 1800, 1600200, true)}${lvl(5, 1800, 2057400, true)}${lvl(6, 1800, 2514600, true)}${lvl(7, 1800, 2971800, true)}${lvl(8, 1800, 3429000, true)}${lvl(9, 1800, 3886200, true)}</p:bodyStyle>
<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${plainLevels}</p:otherStyle>
</p:txStyles></p:sldMaster>`;

function layout(type: string, name: string, shapes: string): string {
  return `${HEAD}<p:sldLayout ${NS} type="${type}" preserve="1"><p:cSld name="${name}"><p:spTree>${GRP}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

const centeredTitle = `<a:bodyPr anchor="b"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r></a:p>`;
const subtitle = `<a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master subtitle style</a:t></a:r></a:p>`;
const footersIn = (start: number) =>
  ph(start, 'Date Placeholder', 'dt', 10, 838200, 6356350, 2743200, 365125) +
  ph(start + 1, 'Footer Placeholder', 'ftr', 11, 4038600, 6356350, 4114800, 365125) +
  ph(start + 2, 'Slide Number Placeholder', 'sldNum', 12, 8610600, 6356350, 2743200, 365125, `<a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`);

const LAYOUTS: [string, string][] = [
  [
    'title',
    layout('title', 'Title Slide', ph(2, 'Title 1', 'ctrTitle', null, 1524000, 1122363, 9144000, 2387600, centeredTitle) + ph(3, 'Subtitle 2', 'subTitle', 1, 1524000, 3602038, 9144000, 1655762, subtitle) + footersIn(4)),
  ],
  [
    'obj',
    layout('obj', 'Title and Content', ph(2, 'Title 1', 'title', null, 838200, 365125, 10515600, 1325563, prompt('Click to edit Master title style')) + ph(3, 'Content Placeholder 2', null, 1, 838200, 1825625, 10515600, 4351338, prompt('Click to edit Master text styles')) + footersIn(4)),
  ],
  [
    'secHead',
    layout('secHead', 'Section Header', ph(2, 'Title 1', 'title', null, 831850, 1709738, 10515600, 2852737, `<a:bodyPr anchor="b"/><a:lstStyle><a:lvl1pPr><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r></a:p>`) + ph(3, 'Text Placeholder 2', 'body', 1, 831850, 4589463, 10515600, 1500187, `<a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p>`) + footersIn(4)),
  ],
  [
    'twoObj',
    layout('twoObj', 'Two Content', ph(2, 'Title 1', 'title', null, 838200, 365125, 10515600, 1325563, prompt('Click to edit Master title style')) + ph(3, 'Content Placeholder 2', null, 1, 838200, 1825625, 5181600, 4351338, prompt('Click to edit Master text styles')) + ph(4, 'Content Placeholder 3', null, 2, 6172200, 1825625, 5181600, 4351338, prompt('Click to edit Master text styles')) + footersIn(5)),
  ],
  ['titleOnly', layout('titleOnly', 'Title Only', ph(2, 'Title 1', 'title', null, 838200, 365125, 10515600, 1325563, prompt('Click to edit Master title style')) + footersIn(3))],
  ['blank', layout('blank', 'Blank', footersIn(2))],
];

const NOTES_MASTER = `${HEAD}<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>
${ph(3, 'Notes Placeholder 2', 'body', 3, 685800, 4400550, 5486400, 3600450, prompt('Click to edit Master text styles'))}
</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:notesStyle>${Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join('')}</p:notesStyle></p:notesMaster>`;

const FIRST_SLIDE = `${HEAD}<p:sld ${NS}><p:cSld><p:spTree>${GRP}
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;

const CT = (slides: string) => `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
${[1, 2, 3, 4, 5, 6].map((i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`).join('')}
<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
${slides}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

/** A fresh, valid .pptx with one empty title slide. */
export function blankDeckBytes(): Uint8Array {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const files: Record<string, string> = {
    '[Content_Types].xml': CT('<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'),
    '_rels/.rels': rels([
      ['rId1', `${RT}/officeDocument`, 'ppt/presentation.xml'],
      ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'],
      ['rId3', `${RT}/extended-properties`, 'docProps/app.xml'],
    ]),
    'docProps/core.xml': `${HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title></dc:title><dc:creator>Greek Office</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
    'docProps/app.xml': `${HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Greek Office · Pinax</Application><PresentationFormat>Widescreen</PresentationFormat></Properties>`,
    'ppt/presentation.xml': PRESENTATION,
    'ppt/_rels/presentation.xml.rels': rels([
      ['rId1', `${RT}/slideMaster`, 'slideMasters/slideMaster1.xml'],
      ['rId2', `${RT}/theme`, 'theme/theme1.xml'],
      ['rId3', `${RT}/notesMaster`, 'notesMasters/notesMaster1.xml'],
      ['rId4', `${RT}/presProps`, 'presProps.xml'],
      ['rId5', `${RT}/viewProps`, 'viewProps.xml'],
      ['rId6', `${RT}/tableStyles`, 'tableStyles.xml'],
      ['rId7', `${RT}/slide`, 'slides/slide1.xml'],
    ]),
    'ppt/presProps.xml': `${HEAD}<p:presentationPr ${NS}/>`,
    'ppt/viewProps.xml': `${HEAD}<p:viewPr ${NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`,
    'ppt/tableStyles.xml': `${HEAD}<a:tblStyleLst xmlns:a="${A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
    'ppt/theme/theme1.xml': THEME('Greek Office'),
    'ppt/theme/theme2.xml': THEME('Greek Office Notes'),
    'ppt/slideMasters/slideMaster1.xml': MASTER,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([
      ...LAYOUTS.map((_, i) => [`rId${i + 1}`, `${RT}/slideLayout`, `../slideLayouts/slideLayout${i + 1}.xml`] as [string, string, string]),
      ['rId7', `${RT}/theme`, '../theme/theme1.xml'],
    ]),
    'ppt/notesMasters/notesMaster1.xml': NOTES_MASTER,
    'ppt/notesMasters/_rels/notesMaster1.xml.rels': rels([['rId1', `${RT}/theme`, '../theme/theme2.xml']]),
    'ppt/slides/slide1.xml': FIRST_SLIDE,
    'ppt/slides/_rels/slide1.xml.rels': rels([['rId1', `${RT}/slideLayout`, '../slideLayouts/slideLayout1.xml']]),
  };
  LAYOUTS.forEach(([, xml], i) => {
    files[`ppt/slideLayouts/slideLayout${i + 1}.xml`] = xml;
    files[`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`] = rels([['rId1', `${RT}/slideMaster`, '../slideMasters/slideMaster1.xml']]);
  });
  const zipped: Record<string, Uint8Array> = {};
  zipped['[Content_Types].xml'] = strToU8(files['[Content_Types].xml']);
  for (const [k, v] of Object.entries(files)) if (k !== '[Content_Types].xml') zipped[k] = strToU8(v);
  return zipSync(zipped, { level: 6 });
}
