let code_part_id // the id of the cusomxml part that holds the code
let css_suffix="" // set by the user with set_css()
const panels=['panel_introduction','panel_examples','panel_code_editor']
const panel_stack=['panel_introduction']
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


function start_me_up(){
  styles.system=document.getElementById("head_style").innerText
  
  panels.push("panel_introduction")
  // fit the editor to the windows on resize
  window.addEventListener('resize', function(event) {
    if(tag("editor-page")){
    tag("editor-page").style.height = (window.innerHeight-86)+"px"
    }
  }, true);
  build_panel("panel_examples")
  build_panel("panel_code_editor")
  init_editor()
  init_examples()
  //done opening, see what is active panel in case auto_exec opened something

  
}

async function get_style(style_name, url, integrate_now){

  if(integrate_now===undefined){
     integrate_now=true
  }

  if(!styles[style_name]){
    styles[style_name]=url
  }

  if(styles[style_name].substr(0,8)==="https://"){
    // the style has not yet been fetched
    console.log("fetching")
    const response = await fetch(styles[style_name])
    const data = await response.text()
    console.log("data",data)
    styles[style_name]=data
    if(integrate_now){
      document.getElementById("head_style").remove()
      document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + styles[style_name] + css_suffix + "</style>")
    }
  }
  
}

function set_style(style_name){

  console.log("at set style", style_name)
  let css_sfx = css_suffix
  if(!style_name){
    style_name="system"
    css_sfx=""

  }

  console.log("styles[style_name].substr(0,8)",styles[style_name].substr(0,8))
  if(styles[style_name].substr(0,8)==="https://"){
    // this style has not been fetched.  Get it now
    console.log("in iff")
    get_style(style_name)
  }

  const style_tag = document.getElementById("head_style")
  if(style_tag.dataset.name!==style_name ){
    // only update the style tag if it is a differnt name
    style_tag.remove()
    document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + styles[style_name] + css_sfx + "</style>")
  }
  
}


function show_examples(){
  console.log(1)
  const panel_name="panel_examples"
  console.log(2)
  set_style()
  console.log(3)
 //console.log(panel_name, tag(panel_name))
  if(!tag(panel_name).innerHTML){
    console.log(4)
    init_examples()
    console.log(5)
  }
  show_panel(panel_name)

}



function build_panel(id){
  const div=document.createElement("div");
  div.id=id
  div.style.display="none"
  if(!panels.includes){
    panels.push(div)
  }
  
document.body.appendChild(div)

}

function show_automations(){
  const panel_name="panel_automations"
 //console.log(1)
  if(!panels.includes){
    build_panel(panel_name)
  }
  

  // get the list of functions
  const code=tag("editor_script_div").firstChild.innerText
  const functions=get_function_names(code).function
  console.log("tag(panel_name)",tag(panel_name))

  const html=['<div onclick="close_canvas(\'panel_automations\')" class="top-corner" style="border: 1px solid black; padding:5px 5px 0 5px;margin:5px 15px 0 0; cursor:pointer"><i class="ms-Icon ms-Icon--ChromeClose"></i></div><h2 style="margin:0 0 0 1rem">Automations built into this workbook.</h2><ol>']

  for(const func of functions){
    let func_value = null
    if(code.includes("async function " + func + "(excel)")){
      //this is an async call to a funtion that interact with the workbook
      func_value=func + "(excel)"
    }else if(code.includes("function " + func + "()")){
      // this is a regular JS function with no params
      func_value=func+"()"
    }
    const function_text = window[func]+''
    if(func_value && function_text.includes("ace.listing:")){ // this is a function we can run directly and it as the comment
      const comment = function_text.split("ace.listing:")[1].split("*/")[0]
     //console.log("comment", comment)
      try{
        const comment_json=JSON.parse(comment)
        let stmt
        if(func_value.includes("(excel)")){
          stmt="Excel.run(" + func_value.split("(")[0] + ")"
        }else{
          stmt=func_value
        }

        html.push('<li onclick="'+stmt+'" style="cursor:pointer"><b>'+comment_json.name+'</b>: '+comment_json.description+'</li>')
       //console.log("html",html)
     }catch(e){
      //console.log("ace.listing was not valid JSON", comment)        
     }
    } 
  }

  if(html.length===1){
    //Did not find any properly configured fucntions
    html[0]="Message to tell the developer how to configure"
  }else{
    html.push("</ul>")  
  }
  open_canvas("panel_listings",html.join(""))
  
}

function show_panel(panel_name){
  if(panels.slice(0, 3).includes(panel_name)){
    set_style()
  }

 // need to see what is open to pop it form the list 
//  for(const panel of panels){
//     if(tag(panel).style.display==='block'){
//       panel_stack.pop()
//     }
//  }
  
 //console.log("trying",panel_name)
  for(const panel of panels){
    if(panel===panel_name){
     //console.log("showing", panel)
      if(tag("selector_"+ panel_name)){
        tag("selector_"+ panel_name).value=panel_name
      }
      tag(panel).style.display="block"  
      panel_stack.push(panel)
    }else{
      console.log(" hiding", panel)
      tag(panel).style.display="none"  
    }
  }
}

function init_editor(){
  const panel_name = "panel_code_editor"
 //console.log("initializing examples", tag(panel_name))
  
  tag(panel_name).appendChild(get_panel_selector(panel_name))
  const editor_container = document.createElement("div")
  editor_container.className="editor-content"
  let code
  //let customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);

 //console.log("code_part_id",code_part_id)
  
  Excel.run(async (excel)=>{
    let customXmlPart = excel.workbook.customXmlParts.getItem(code_part_id);
    const xmlBlob = customXmlPart.getXml();
    excel.sync().then(function() { 
      //console.log("xmlBlob",xmlBlob.value)
      code=atob(xmlBlob.value.substr(58, xmlBlob.value.length-77 ))
     //console.log(code)
      //console.log(atob(code))

      // const parser = new DOMParser();
      // const doc = parser.parseFromString(xmlBlob.value, "application/xml");
      ////console.log("doc",doc)

      let div = document.createElement("div");
      div.style.verticalAlign="middle"
      div.style.verticalAlign.padding=".2rem"
      div.innerHTML = ' <button class="button-14" id="code-saver" onclick="update_editor_script()">Save</button> <button class="button-14" id="code-runner" onclick="code_runner(tag(\'function-names\').value)">Run</button> <select id="function-names"></select>';
      div.style.height="40px"
      div.style.paddingLeft=".5rem"
      div.style.backgroundColor = "DimGray";
      div.id="editor-bar"
      editor_container.appendChild(div);


      const box = document.createElement("div");
      box.id = "editor-page";
      box.style.width = "100%";
      box.style.height = (window.innerHeight-86)+"px"
      box.style.display = "inline-block";
      box.style.position = "relative";

     //console.log("=======================================")
     //console.log(window.innerHeight);
     //console.log("document",document.body.clientHeight);
     //console.log("scr",tag("panel_code_editor").Height);


      div = document.createElement("div");
      div.id = "editor-content";
      div.dataset.edited=false
      div.innerHTML = code.toHtmlEntities();

      box.appendChild(div);


      editor_container.appendChild(box);

      //        elem.innerHTML = '<pre id="pre' + id + '">' + gist.script.content.split("<").join("&lt;") + '</pre>'

      const scriptPromise = new Promise((resolve, reject) => {
        const script = document.createElement("script");
        document.body.appendChild(script);
        script.onload = resolve;
        script.onerror = reject;
        script.async = true;
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.13/ace.js";
      });

      scriptPromise.then(() => {
        var editor = ace.edit("editor-content");
        editor.getSession().setUseWorker(false);
        editor.on("blur", function () {
          window.event.target.parentElement.dataset.edited=true
        });
        editor.setTheme("ace/theme/solarized_light");
        editor.setOptions({fontSize: "14pt"});
        editor.session.setMode("ace/mode/javascript");
      });

      editor_container.style.display = "block";

      div = document.createElement("div");
      div.id = "editor_script_div";      
      const script = document.createElement("script");
      script.innerHTML = code;
      div.appendChild(script);
      editor_container.appendChild(div);

      tag("panel_code_editor").appendChild(editor_container)

      load_function_names_select(code)

      // AutoExecutable function
     //console.log("about to autoexec")
      try{
       //console.log("in try")
        auto_exec()
      }catch(e){
       ;console.log("catch",e)
      }  


    })
  })
}

function code_runner(script_name){
  if (tag("editor-content").dataset.edited==="true"){
    update_editor_script()
  }
  if(script_name.includes("(excel)")){
    setTimeout("Excel.run(" + script_name.split("(")[0] + ")", 0) //run the function
  }else{
    setTimeout(script_name, 0) //run the function
  }
}

function init_examples(){
  const panel_name="panel_examples"
  const panel=tag(panel_name)
 //console.log("initializing examples")
  panel.appendChild(get_panel_selector(panel_name))
  const div = document.createElement("div")
  div.className="content"
  div.id="e_content"
  panel.appendChild(div)

//console.log("about to fetch")

fetch("https://thegove.github.io/VBA-samples/excel-js_snippets.yaml?" + new Date())
.then((response) => response.text())
.then((yaml_data) => {
  const html = [];
  let i = 0;
  const data = jsyaml.load(yaml_data);
  for (const group of data) {
    html.push("<h2>" + group.name + "</h2>");
    for (const gist of group.snips) {
      html.push(
        '<p><a title="Copy Gist URL for import" onclick="copy(\'' +
        gist.URL +
        '\')"><img src="https://thegove.github.io/VBA-samples/link.png" width="18px" class="link"></a><a class="link" title="Show Code" onclick="show_example(' +
        ++i +
        ",'" +
        gist.URL +
        "')\">" +
        gist.name +
        "</a> " +
        gist.description +
        '</p><div id="page' +
        i +
        '"></div>'
      );
    }
  }
  document.getElementById("e_content").innerHTML = html.join("");
});


}

function show_example(id, url) {
  // place the code in an editable box for user to see and play with
  const elem = tag("page" + id);
  //console.log(id + id);
  if (elem.innerHTML === "") {
    fetch(url.replace("gist.github.com", "gist.githubusercontent.com") + "/raw/?" + new Date())
      .then((response) => response.text())
      .then((data) => {
        const gist = jsyaml.load(data);
        //console.log(gist)

        let div = document.createElement("div");
        div.innerHTML = gist.template.content;
        div.style.marginBottom = "1rem";
        elem.appendChild(div);

        const box = document.createElement("div");
        box.id = "page" + id;
        box.style.width = "100%";
        box.style.height = "170px";
        box.style.display = "inline-block";
        box.style.position = "relative";

        div = document.createElement("div");
        div.id = "editor" + id;
        div.innerHTML = gist.script.content;

        box.appendChild(div);
        elem.appendChild(box);

        //        elem.innerHTML = '<pre id="pre' + id + '">' + gist.script.content.split("<").join("&lt;") + '</pre>'

        const scriptPromise = new Promise((resolve, reject) => {
          const script = document.createElement("script");
          document.body.appendChild(script);
          script.onload = resolve;
          script.onerror = reject;
          script.async = true;
          script.src = "https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.13/ace.js";
        });

        scriptPromise.then(() => {
          var editor = ace.edit("editor" + id);
          editor.on("blur", function () {
            update_script(id);
          });
          editor.setTheme("ace/theme/solarized_light");
          //editor.session.$worker.send("changeOptions", [{ asi: true }]);
          editor.setOptions({fontSize: "14pt"});
          editor.getSession().setUseWorker(false);
          editor.session.setMode("ace/mode/javascript");
        });

        elem.style.display = "block";

        div = document.createElement("div");
        div.id = "script" + id;
        const script = document.createElement("script");
        script.innerHTML = gist.script.content;
        div.appendChild(script);
        elem.appendChild(div);
      })
      .catch((error) => {
       //console.log(error);
      });
  } else {
    if (elem.style.display === "block") {
      elem.style.display = "none";
    } else {
      elem.style.display = "block";
    }
  }
}

function copy(text, out) {
  navigator.clipboard
    .writeText(text)
    .then(() => {
      if (out) {
        status('choose "Import" from the Script Lab code editor main menu', out);
      }
    })
    .catch(() => {
      //console.log("Failed to copy text.");
    });
}

function update_script(id) {
  // read the script for an ace editor and write it to the DOM
  //console.log("script" + id);
  script_div = document.getElementById("script" + id);
  script_div.innerHTML = "";
  const editor = ace.edit("editor" + id);
  const script = document.createElement("script");
  script.innerHTML = editor.getValue();
  script_div.appendChild(script);
}


function update_editor_script() {
  // read the script for an ace editor and write it to the DOM
  // also saves the module to the custom properties
  script_div = document.getElementById("editor_script_div");
  script_div.innerHTML = "";
  
  const code = ace.edit("editor-content").getValue();
 //console.log(code)
  const script = document.createElement("script");
  script.innerHTML = code
  script_div.appendChild(script);
  
  load_function_names_select(code)

  // save to workbook
  Excel.run(async (excel)=>{
      excel.sync().then(function() { 
      const customXmlPart = excel.workbook.customXmlParts.getItem(code_part_id);
      customXmlPart.setXml("<Modules xmlns='http://schemas.gove.net/code/1.0'><Module>" + btoa(code) + "</Module></Modules>");
      const xmlBlob = customXmlPart.getXml();
      excel.sync();
    })
  })

}



function load_function_names_select(code){// reads teh function names from the code and puts them in the function name select
  const selectElement=tag("function-names")
  const selected_script = selectElement.value
  

  while(selectElement.options.length>0) {
     selectElement.remove(0);
  }

  for(const func of get_function_names(code).function){
    let func_value = null
    if(code.includes("async function " + func + "(excel)")){
      //this is an async call to a funtion that interact with the workbook
      func_value=func + "(excel)"
    }else if(code.includes("function " + func + "()")){
      // this is a regular JS function with no params
      func_value=func+"()"
    }

    if(func_value){ // this is a function we can run directly
      const option = document.createElement("option")
      if(func_value===selected_script){option.selected='selected'}
      option.text = func
      option.value = func_value
      selectElement.add(option)    
    } 

  }
}

function get_function_names(text) {  // gets the names of functions from a block of JS
  text = text.replace(/\/\/.*?\r?\n/g, "")                                 // first, remove line comments
             .replace(/\r?\n/g, " ")                                       // then remove new lines (replace them with spaces to not break the structure)
             .replace(/\/\*.*?\*\//g, "");                                 // then remove block comments
             
  // PART 1: Match functions declared using: var * = function 
  var varFuncs      = (text.match(/[$A-Z_][0-9A-Z_$]*\s*=\s*function[( ]/gi) || []) // match any valid function name that comes before \s*=\s*function
                           .map(function(tex) {                            // then extract only the function names from the matches
                             return tex.match(/^[$A-Z_][0-9A-Z_$]*/i)[0];
                           });

  // PART 2: Match functions declared using: function * 
  var functionFuncs = (text.match(/function\s+[^(]+/g) || [])              // match anything that comes after function and before (
                           .map(function(tex) {                            // then extarct only the names from the matches
                             return tex.match(/[$A-Z_][0-9A-Z_$]*$/i)[0];
                           });
  return {
    var: varFuncs,
    function: functionFuncs
  };
} 

function get_panel_selector(panel){
  const sel = document.createElement("select")
  const local_panels=["Introduction", "Examples", "Code Editor"]
  sel.className="panel-selector"
  for (let i=0; i<local_panels.length; i++) {
    var option = document.createElement("option");
    option.value = "panel_" + local_panels[i].toLocaleLowerCase().split(" ").join("_");
   //console.log("-->", option.value)
    option.text = local_panels[i];
    option.className="panel-selector-option"
    sel.appendChild(option);
  }
 //console.log(panel)
  sel.value=panel
  sel.style.height="40px"
  sel.id="selector_"+panel
  sel.onchange = select_page
  return sel
}


function select_page(){
  show_panel(window.event.target.value)
}



/**
 * Convert a string to HTML entities
 */
 String.prototype.toHtmlEntities = function() {
  return this.replace(/./gm, function(s) {
      // return "&#" + s.charCodeAt(0) + ";";
      return (s.match(/[a-z0-9\s]+/i)) ? s : "&#" + s.charCodeAt(0) + ";";
  });
};

/**
* Create string from HTML entities
*/
String.fromHtmlEntities = function(string) {
  return (string+"").replace(/&#\d+;/gm,function(s) {
      return String.fromCharCode(s.match(/\d+/gm)[0]);
  })
};


