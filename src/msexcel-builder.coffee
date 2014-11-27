###
  MS Excel 2007 Creater v0.0.1
  Author : chuanyi.zheng@gmail.com
  Edit-Author: fungzhiwen@126.com
  History: 2012/11/07 first created
###

fs  = require 'fs'
path = require 'path'
exec = require 'child_process'
xml = require 'xmlbuilder'
require 'node-zip'
existsSync = fs.existsSync || path.existsSync
templateXLSX = "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";

tool = 
  i2a : (i) ->
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt(i-1)

  copy : (origin, target) ->
  	if existsSync(origin)
      fs.mkdirSync(target, 0o755) if not existsSync(target)
      files = fs.readdirSync(origin)
      if files
        for f in files
          oCur = origin + '/' + f
          tCur = target + '/' + f
          s = fs.statSync(oCur)
          if s.isFile()
            fs.writeFileSync(tCur,fs.readFileSync(oCur,''),'')
          else
            if s.isDirectory()
              tool.copy oCur, tCur

opt = 
  tmpl_path : __dirname

class ContentTypes
  constructor: (@book)->

  toxml:()->
    types = xml.create('Types',{version:'1.0',encoding:'UTF-8',standalone:true})
    types.att('xmlns','http://schemas.openxmlformats.org/package/2006/content-types')
    types.ele('Override',{PartName:'/xl/theme/theme1.xml',ContentType:'application/vnd.openxmlformats-officedocument.theme+xml'})
    types.ele('Override',{PartName:'/xl/styles.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
    types.ele('Default',{Extension:'rels',ContentType:'application/vnd.openxmlformats-package.relationships+xml'})
    types.ele('Default',{Extension:'xml',ContentType:'application/xml'})
    # add psdmcp by fungzw
    types.ele('Default',{Extension:'psmdcp',ContentType:'application/vnd.openxmlformats-package.core-properties+xml'})
    types.ele('Override',{PartName:'/xl/workbook.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    types.ele('Override',{PartName:'/docProps/app.xml',ContentType:'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
    for i in [1..@book.sheets.length]
      types.ele('Override',{PartName:'/xl/worksheets/sheet'+i+'.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
    types.ele('Override',{PartName:'/xl/sharedStrings.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
    types.ele('Override',{PartName:'/docProps/core.xml',ContentType:'application/vnd.openxmlformats-package.core-properties+xml'})
    return types.end()

class DocPropsApp
  constructor: (@book)->

  toxml: ()->
    props = xml.create('Properties',{version:'1.0',encoding:'UTF-8',standalone:true})
    props.att('xmlns','http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
    props.att('xmlns:vt','http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
    props.ele('Application','Microsoft Excel')
    props.ele('DocSecurity','0')
    props.ele('ScaleCrop','false')
    tmp = props.ele('HeadingPairs').ele('vt:vector',{size:2,baseType:'variant'})
    tmp.ele('vt:variant').ele('vt:lpstr','工作表')
    tmp.ele('vt:variant').ele('vt:i4',''+@book.sheets.length)
    tmp = props.ele('TitlesOfParts').ele('vt:vector',{size:@book.sheets.length,baseType:'lpstr'})
    for i in [1..@book.sheets.length]
      tmp.ele('vt:lpstr',@book.sheets[i-1].name)
    props.ele('Company')
    props.ele('LinksUpToDate','false')
    props.ele('SharedDoc','false')  
    props.ele('HyperlinksChanged','false')  
    props.ele('AppVersion','12.0000') 
    return props.end()

class XlWorkbook
  constructor: (@book)->

  toxml: ()->
    wb = xml.create('workbook',{version:'1.0',encoding:'UTF-8',standalone:true})
    wb.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    wb.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    wb.ele('fileVersion ',{appName:'xl',lastEdited:'4',lowestEdited:'4',rupBuild:'4505'})
    wb.ele('workbookPr',{filterPrivacy:'1',defaultThemeVersion:'124226'}) 
    wb.ele('bookViews').ele('workbookView ',{xWindow:'0',yWindow:'90',windowWidth:'19200',windowHeight:'11640'})
    tmp = wb.ele('sheets')
    for i in [1..@book.sheets.length]
      tmp.ele('sheet',{name:@book.sheets[i-1].name,sheetId:''+i,'r:id':'rId'+i})
    wb.ele('calcPr',{calcId:'124519'})
    return wb.end()

class XlRels
  constructor: (@book)->
  
  toxml: ()->
    rs = xml.create('Relationships',{version:'1.0',encoding:'UTF-8',standalone:true})
    rs.att('xmlns','http://schemas.openxmlformats.org/package/2006/relationships')
    for i in [1..@book.sheets.length]
      rs.ele('Relationship',{Id:'rId'+i,Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',Target:'worksheets/sheet'+i+'.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+1),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',Target:'theme/theme1.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+2),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',Target:'styles.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+3),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',Target:'sharedStrings.xml'})
    return rs.end()

class SharedStrings
  constructor: ()->
    @cache = {}
    @arr = []

  str2id: (s)->
    id = @cache[s]
    if id
      return id
    else
      @arr.push s
      @cache[s] = @arr.length
      return @arr.length

  toxml: ()->
    sst = xml.create('sst',{version:'1.0',encoding:'UTF-8',standalone:true})
    sst.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    sst.att('count',''+@arr.length)
    sst.att('uniqueCount',''+@arr.length)
    for i in [0...@arr.length]
      si = sst.ele('si')
      si.ele('t',@arr[i])
      si.ele('phoneticPr',{fontId:1,type:'noConversion'})
    return sst.end()

class Sheet
  constructor: (@book, @name, @cols, @rows) ->
    @data = {}
    for i in [1..@rows]
      @data[i] = {}
      for j in [1..@cols]
        @data[i][j] = {v:0}
    @merges = []
    @col_wd = []
    @row_ht = {}
    @styles = {}

  set: (col, row, str) ->
    @data[row][col].v = @book.ss.str2id(''+str) if str? and str isnt ''

  merge: (from_cell, to_cell) ->
    @merges.push({from:from_cell, to:to_cell})

  width: (col, wd) ->
    @col_wd.push {c:col,cw:wd}

  height: (row, ht) ->
    @row_ht[row] = ht

  font: (col, row, font_s)->
    @styles['font_'+col+'_'+row] = @book.st.font2id(font_s)

  fill: (col, row, fill_s)-> 
    @styles['fill_'+col+'_'+row] = @book.st.fill2id(fill_s)

  border: (col, row, bder_s)->
    @styles['bder_'+col+'_'+row] = @book.st.bder2id(bder_s)

  align: (col, row, align_s)->
    @styles['algn_'+col+'_'+row] = align_s

  valign: (col, row, valign_s)->
    @styles['valgn_'+col+'_'+row] = valign_s

  rotate: (col, row, textRotation)->
    @styles['rotate_'+col+'_'+row] = textRotation

  wrap: (col, row, wrap_s)->
    @styles['wrap_'+col+'_'+row] = wrap_s

  style_id: (col, row) ->
    inx = '_'+col+'_'+row
    style = {font_id:@styles['font'+inx],fill_id:@styles['fill'+inx],bder_id:@styles['bder'+inx],align:@styles['algn'+inx],valign:@styles['valgn'+inx],rotate:@styles['rotate'+inx],wrap:@styles['wrap'+inx]}
    id = @book.st.style2id(style)
    return id

  toxml: () ->
    ws = xml.create('worksheet',{version:'1.0',encoding:'UTF-8',standalone:true})
    ws.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    ws.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ws.ele('dimension',{ref:'A1'})
    ws.ele('sheetViews').ele('sheetView',{workbookViewId:'0'})
    ws.ele('sheetFormatPr',{defaultRowHeight:'13.5'})
    if @col_wd.length > 0
      cols = ws.ele('cols')
      for cw in @col_wd
        cols.ele('col',{min:''+cw.c,max:''+cw.c,width:cw.cw,customWidth:'1'})
    sd = ws.ele('sheetData')
    for i in [1..@rows]
      r = sd.ele('row',{r:''+i,spans:'1:'+@cols})
      ht = @row_ht[i]
      if ht
        r.att('ht',ht)
        r.att('customHeight','1')
      for j in [1..@cols]
        ix = @data[i][j]
        sid = @style_id(j, i)
        if (ix.v isnt 0) or (sid isnt 1)
          c = r.ele('c',{r:''+tool.i2a(j)+i})
          c.att('s',''+(sid-1)) if sid isnt 1
          if ix.v isnt 0
            c.att('t','s')
            c.ele('v',''+(ix.v-1))
    if @merges.length > 0
      mc = ws.ele('mergeCells',{count:@merges.length})
      for m in @merges
        mc.ele('mergeCell',{ref:(''+tool.i2a(m.from.col)+m.from.row+':'+tool.i2a(m.to.col)+m.to.row)})
    ws.ele('phoneticPr',{fontId:'1',type:'noConversion'})
    ws.ele('pageMargins',{left:'0.7',right:'0.7',top:'0.75',bottom:'0.75',header:'0.3',footer:'0.3'})
    ws.ele('pageSetup',{paperSize:'9',orientation:'portrait',horizontalDpi:'200',verticalDpi:'200'})
    return ws.end()

class Style
  constructor: (@book)->
    @cache = {}
    @mfonts = []  # font style
    @mfills = []  # fill style
    @mbders = []  # border style
    @mstyle = []  # cell style<ref-font,ref-fill,ref-border,align>
    @with_default()

  with_default:()->
    @def_font_id = @font2id(null)
    @def_fill_id = @fill2id(null)
    @def_bder_id = @bder2id(null)
    @def_align = '-'
    @def_valign = '-'
    @def_rotate = '-'
    @def_wrap = '-'
    @def_style_id = @style2id({font_id:@def_font_id,fill_id:@def_fill_id,bder_id:@def_bder_id,align:@def_align,valign:@def_valign,rotate:@def_rotate})

  font2id: (font)->
    font or= {}
    font.bold or= '-'
    font.iter or= '-'
    font.sz or= '11'
    font.color or= '-'
    font.name or= '宋体'
    font.scheme or='minor'
    font.family or= '2'
    k = 'font_'+font.bold+font.iter+font.sz+font.color+font.name+font.scheme+font.family
    id = @cache[k]
    if id
      return id
    else
      @mfonts.push font
      @cache[k] = @mfonts.length
      return @mfonts.length

  fill2id: (fill)->
    fill or= {}
    fill.type or= 'none'
    fill.bgColor or= '-'
    fill.fgColor or= '-'
    k = 'fill_' + fill.type + fill.bgColor + fill.fgColor
    id = @cache[k]
    if id
      return id
    else
      @mfills.push fill
      @cache[k] = @mfills.length
      return @mfills.length

  bder2id: (bder)->
    bder or= {}
    bder.left or= '-'
    bder.right or= '-'
    bder.top or= '-'
    bder.bottom or= '-'
    k = 'bder_'+bder.left+'_'+bder.right+'_'+bder.top+'_'+bder.bottom
    id = @cache[k]
    if id
      return id
    else
      @mbders.push bder
      @cache[k] = @mbders.length
      return @mbders.length

  style2id:(style)->
    style.align or= @def_align
    style.valign or= @def_valign
    style.rotate or= @def_rotate
    style.wrap or= @def_wrap
    style.font_id or= @def_font_id
    style.fill_id or= @def_fill_id
    style.bder_id or= @def_bder_id
    k = 's_' + style.font_id + '_' + style.fill_id + '_' + style.bder_id + '_' + style.align + '_' + style.valign + '_' + style.wrap + '_' + style.rotate
    id = @cache[k]
    if id
      return id
    else
      @mstyle.push style
      @cache[k] = @mstyle.length
      return @mstyle.length

  toxml: ()->
    ss = xml.create('styleSheet',{version:'1.0',encoding:'UTF-8',standalone:true})
    ss.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    fonts = ss.ele('fonts',{count:@mfonts.length})
    for o in @mfonts
      e = fonts.ele('font')
      e.ele('b') if o.bold isnt '-'
      e.ele('i') if o.iter isnt '-'
      e.ele('sz',{val:o.sz})
      # edit by fungzw(color has two different part,theme or rgb)
      if o.color isnt '-'
        if o.color.rgb isnt '-'
          e.ele('color',{rgb:o.color.rgb})
        else
          e.ele('color',{theme:o.color.rgb})
      #e.ele('color',{theme:o.color}) if o.color isnt '-'
      e.ele('name',{val:o.name})
      e.ele('family',{val:o.family})
      e.ele('charset',{val:'134'})
      e.ele('scheme',{val:'minor'}) if o.scheme isnt '-'
    fills = ss.ele('fills',{count:@mfills.length})
    for o in @mfills
      e = fills.ele('fill')
      es = e.ele('patternFill',{patternType:o.type})
      es.ele('fgColor',{theme:'8',tint:'0.79998168889431442'}) if o.fgColor isnt '-'
      es.ele('bgColor',{indexed:o.bgColor}) if o.bgColor isnt '-'
    bders = ss.ele('borders',{count:@mbders.length})
    for o in @mbders
      e = bders.ele('border')
      if o.left isnt '-' then e.ele('left',{style:o.left}).ele('color',{auto:'1'}) else e.ele('left')
      if o.right isnt '-' then e.ele('right',{style:o.right}).ele('color',{auto:'1'}) else e.ele('right')
      if o.top isnt '-' then e.ele('top',{style:o.top}).ele('color',{auto:'1'}) else e.ele('top')
      if o.bottom isnt '-' then e.ele('bottom',{style:o.bottom}).ele('color',{auto:'1'}) else e.ele('bottom')
      e.ele('diagonal')
    ss.ele('cellStyleXfs',{count:'1'}).ele('xf',{numFmtId:'0',fontId:'0',fillId:'0',borderId:'0'}).ele('alignment',{vertical:'center'})
    cs = ss.ele('cellXfs',{count:@mstyle.length})
    for o in @mstyle
      e = cs.ele('xf',{numFmtId:'0',fontId:(o.font_id-1),fillId:(o.fill_id-1),borderId:(o.bder_id-1),xfId:'0'})
      e.att('applyFont','1') if o.font_id isnt 1
      e.att('applyFill','1') if o.fill_id isnt 1
      e.att('applyBorder','1') if o.bder_id isnt 1
      if o.align isnt '-' or o.valign isnt '-' or o.wrap isnt '-'
        e.att('applyAlignment','1')
        ea = e.ele('alignment',{textRotation:(if o.rotate is '-' then '0' else o.rotate),horizontal:(if o.align is '-' then 'left' else o.align), vertical:(if o.valign is '-' then 'top' else o.valign)})
        ea.att('wrapText','1') if o.wrap isnt '-'
    ss.ele('cellStyles',{count:'1'}).ele('cellStyle',{name:'常规',xfId:'0',builtinId:'0'})
    ss.ele('dxfs',{count:'0'})
    ss.ele('tableStyles',{count:'0',defaultTableStyle:'TableStyleMedium9',defaultPivotStyle:'PivotStyleLight16'})
    return ss.end()

class Workbook
  constructor: (@fpath, @fname, @isWeb) ->
    # if is web,don't create temp folder
    unless @isWeb
      @id = ''+parseInt(Math.random()*9999999)
      # create temp folder & copy template data
      target = @fpath + '/' + @id + '/'
      fs.rmdirSync(target) if existsSync(target)
      tool.copy (opt.tmpl_path + '/tmpl'),target
    # init
    @sheets = []
    @ss = new SharedStrings
    @ct = new ContentTypes(@)
    @da = new DocPropsApp(@)
    @wb = new XlWorkbook(@)
    @re = new XlRels(@)
    @st = new Style(@)

  createSheet: (name, cols, rows) ->
    sheet = new Sheet(@,name,cols,rows)
    @sheets.push sheet
    return sheet

  save: (cb) =>
    target = @fpath + '/' + @id
    # 1 - build [Content_Types].xml
    fs.writeFileSync(target+'/[Content_Types].xml',@ct.toxml(),'utf8')
    # 2 - build docProps/app.xml
    fs.writeFileSync(target+'/docProps/app.xml',@da.toxml(),'utf8')
    # 3 - build xl/workbook.xml
    fs.writeFileSync(target+'/xl/workbook.xml',@wb.toxml(),'utf8')
    # 4 - build xl/sharedStrings.xml
    fs.writeFileSync(target+'/xl/sharedStrings.xml',@ss.toxml(),'utf8')
    # 5 - build xl/_rels/workbook.xml.rels
    fs.writeFileSync(target+'/xl/_rels/workbook.xml.rels',@re.toxml(),'utf8')
    # 6 - build xl/worksheets/sheet(1-N).xml
    for i in [0...@sheets.length]
      fs.writeFileSync(target+'/xl/worksheets/sheet'+(i+1)+'.xml',@sheets[i].toxml(),'utf8')
    # 7 - build xl/styles.xml
    fs.writeFileSync(target+'/xl/styles.xml',@st.toxml(),'utf8')    
    # 8 - compress temp folder to target file
    args = ' a -tzip "' + @fpath + '/' + @fname + '" "*"'
    opts = {cwd:target}
    exec.exec '"'+opt.tmpl_path+'/tool/7za.exe"' + args, opts, (err,stdout,stderr)->
      # 9 - delete temp folder
      exec.exec 'rmdir "' + target + '" /q /s',()->
        cb not err

  # edit by fungzw(not save local file)
  saveToWeb: () =>
    xlsx = new JSZip(templateXLSX, { base64: true, checkCRC32: false })

    # 1 - build [Content_Types].xml
    xlsx.file('[Content_Types].xml', @ct.toxml())
    # 2 - build docProps/app.xml
    xlsx.file('docProps/app.xml', @da.toxml())
    # 3 - build xl/workbook.xml
    xlsx.file("xl/workbook.xml", @wb.toxml())
    # 4 - build xl/sharedStrings.xml
    xlsx.file('xl/sharedStrings.xml', @ss.toxml())
    # 5 - build xl/_rels/workbook.xml.rels
    xlsx.file('xl/_rels/workbook.xml.rels', @re.toxml())
    # 6 - build xl/worksheets/sheet(1-N).xml
    for i in [0...@sheets.length]
      xlsx.file("xl/worksheets/sheet" + (i + 1) + '.xml',@sheets[i].toxml())
    # 7 - build xl/styles.xml
    xlsx.file('xl/styles.xml', @st.toxml())

    results = xlsx.generate({ base64: false, compression: "DEFLATE" })

  cancel: () ->
    # delete temp folder
    fs.rmdirSync target

module.exports = 
  createWorkbook: (fpath, fname, isWeb)->
    return new Workbook(fpath, fname, isWeb)
