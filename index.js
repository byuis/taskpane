let code_part_id // the id of the cusomxml part that holds the code




function start_me_up(){
  console.log ("starting")
  window.addEventListener('resize', function(event) {
    if(tag("editor-page")){
    tag("editor-page").style.height = (window.innerHeight-86)+"px"
    }
  }, true);

}


function show_panel(show){
  console.log("showing", show)
  for(let i=1;i<4;i++){
    tag("panel"+i).style.display="none"  
  }
  tag("panel"+show).style.display="block"  

  console.log("showing panel"+show)
  console.log(tag("panel"+show).id)
  if(!tag("panel"+show).innerHTML){
    //Initialize the panel
    console.log("==================================================", show, show===3)
    switch(show.toString()){
      case "2":
        init_examples()
        break;
      case "3":
        init_editor()
        break;
      default:
    }
  }
  tag("selector"+show).value=show
}

function init_editor(){
  console.log("initializing examples")
  tag("panel3").appendChild(get_panel_selector(3))
  const editor_container = document.createElement("div")
  editor_container.className="editor-content"
  let code
  //let customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);

  console.log("code_part_id",code_part_id)
  
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
      // console.log("doc",doc)

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

      console.log("=======================================")
      console.log(window.innerHeight);
      console.log("document",document.body.clientHeight);
      console.log("scr",tag("panel3").Height);


      div = document.createElement("div");
      div.id = "editor-content";
      div.dataset.edited=false
      div.innerHTML = code;

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
      


      tag("panel3").appendChild(editor_container)

      load_function_names_select(code)


    })
  })
}

function code_runner(script_name){
  
  console.log('tag("editor-content").dataset.edited',tag("editor-content").dataset.edited)
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
  const panel=tag("panel2")
  console.log("initializing examples")
  panel.appendChild(get_panel_selector(2))
  const div = document.createElement("div")
  div.className="content"
  div.id="e_content"
  panel.appendChild(div)

console.log("about to fetch")

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
        console.log(error);
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
  const editor = ace.edit("editor-content");
  const code = editor.getValue();
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

  for(const func of getNames(code).function){
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

function getNames(text) {  // gets the names of functions from a block of JS
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

function get_panel_selector(value){
  const sel = document.createElement("select")
  const panels=['Introduction','Examples','Editor']
  sel.className="panel-selector"
  for (let i=0; i<panels.length; i++) {
    var option = document.createElement("option");
    option.value = i+1;
    option.text = panels[i];
    sel.appendChild(option);
  }
  sel.value=value
  sel.style.height="40px"
  sel.id="selector"+value
  sel.onchange = selectPage
  return sel
}


function selectPage(){
  console.log ("selecting", window.event.target.value)
  show_panel(window.event.target.value)
}



function tag(id){
  return document.getElementById(id)
}
