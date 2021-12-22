let global_settings={
  system:{
    examples_url:"https://gist.github.com/theGove/989fc9443c232391ecfa83089e24e791"
  },
  ace_options:{
    selectionStyle:"line",
    highlightActiveLine:true,
    highlightSelectedWord:true,
    readOnly:false,
    copyWithEmptySelection:false,
    cursorStyle:"ace",
    mergeUndoDeltas:true,
    behavioursEnabled:true,
    wrapBehavioursEnabled:true,
    enableAutoIndent:true,
    showLineNumbers:true,
    hScrollBarAlwaysVisible:false,
    vScrollBarAlwaysVisible:false,
    highlightGutterLine:true,
    animatedScroll:false,
    showInvisibles:false,
    showPrintMargin:false,
    printMarginColumn:80,
    printMargin:80,
    fadeFoldWidgets:false,
    showFoldWidgets:true,
    displayIndentGuides:true,
    showGutter:true,
    fontSize:"14pt",
    scrollPastEnd:0,
    theme:"ace/theme/solarized_light",
    maxPixelHeight:0,
    useTextareaForIME:true,
    scrollSpeed:2,
    dragDelay:0,
    dragEnabled:true,
    focusTimeout:0,
    tooltipFollowsMouse:true,
    firstLineNumber:1,
    overwrite:false,
    newLineMode:"auto",
    useWorker:false,
    useSoftTabs:true,
    navigateWithinSoftTabs:false,
    tabSize:4,
    wrap:"off",
    indentedSoftWrap:true,
    foldStyle:"markbegin",
    mode:"ace/mode/javascript",
    enableMultiselect:true,
    enableBlockSelect:true,
    tabSize:2,
    useSoftTabs:true
  }
}
let ace_system_module_id=null
let hide_ace_system_module = true
let code_module_ids=[] // the set of code modules read from the workbook's custom xml properties
let css_suffix="" // set by the user with set_css()
const panels=['panel_introduction','panel_examples']
const panel_labels=["Introduction", "Examples", "Output"]
const code_panels=[]
const panel_stack=['panel_introduction']

let global_wrap="off"
let global_theme="ace/theme/solarized_light"


const styles={
  system:null,
  none:"/*No Theme CSS Used*/",
  mvp:"https://cdnjs.cloudflare.com/ajax/libs/mvp.css/1.8.0/mvp.css",
  marx:"https://cdnjs.cloudflare.com/ajax/libs/marx/4.0.0/marx.min.css",
  water:"https://cdn.jsdelivr.net/npm/water.css@2/out/water.css",
  "dark water":"https://cdn.jsdelivr.net/npm/water.css@2/out/dark.css",
  sajura:"https://unpkg.com/sakura.css/css/sakura.css",
  tacit:"https://cdn.jsdelivr.net/gh/yegor256/tacit@gh-pages/tacit-css-1.5.5.min.css",
  pure:"https://unpkg.com/purecss@2.0.6/build/pure-min.css",
  picnic:"https://cdn.jsdelivr.net/npm/picnic",
  wing:"https://unpkg.com/wingcss",
  chota:"https://unpkg.com/chota@latest",
  bootstrap:"https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css",
}