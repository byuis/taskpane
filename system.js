function start_me_up(){
  //console.log("at start_me_up")
  styles.system=tag("head_style").innerText
  panels.push("panel_introduction")

  // add event listner to "add code module" input
  tag("new-module-name").addEventListener("keyup", function(event) {
      event.preventDefault();
      if (event.key === 'Enter') {
          tag("module-add-button").click();
      }
  });

  // add event listner to "import module" input
  tag("gist-url").addEventListener("keyup", function(event) {
    event.preventDefault();
    if (event.key === 'Enter') {
        tag("module-import-button").click();
    }
  });


  // fit the editor to the windows on resize
  window.addEventListener('resize', function(event) {
    //console.log("hi")
    for(const panel_name of code_panels){
      tag(panel_name + "_editor-page").style.height = editor_height()
    }
  }, true);
  init_examples()
  
  // ---------------- Initializing Code Editors -----------------------------
  if(code_module_ids.length>0){// show the button to view code modules
    show_element("open-editor")
  }
  //console.log("at init_code_editors       code_module_ids",code_module_ids)
  Excel.run(async (excel)=>{
    const parser = new DOMParser();
    for(const code_module_id of code_module_ids){
      const xmlpart=excel.workbook.customXmlParts.getItem(code_module_id)
      const xmlBlob = xmlpart.getXml();
      await excel.sync()
      //console.log("blob", xmlBlob)
      const doc = parser.parseFromString(xmlBlob.value, "application/xml");
      const module_name=doc.getElementsByTagName("name")[0].textContent
      const module_code=atob(doc.getElementsByTagName("code")[0].textContent)
      const settings=atob(doc.getElementsByTagName("settings")[0].textContent)
      const options=atob(doc.getElementsByTagName("options")[0].textContent)
      //console.log("settings", settings)
      //console.log("settings2", JSON.parse(settings))
      //console.log("options", options)
      //console.log("options-parsed", JSON.parse(options))
      if(module_name==="ACE System" && hide_ace_system_module){
        // just need to load the code  
        const script = document.createElement("script");
        script.innerHTML = module_code
        script.id="ace_system_script"
        document.body.appendChild(script)
        ace_system_module_id=code_module_id
        tag("ace_system_script").remove()// once the script has been loaded, there is no reason to leave it in the html
      }else{
        add_code_editor(module_name, module_code,code_module_id, JSON.parse(settings), JSON.parse(options))        
      }
    }
    if(ace_system_module_id){
      // the system settings module has been loaded
      // waiting to give the ace editor time to load
      setTimeout(function() {
        try{
        apply_ace_system_settings()
        }catch(e){
          setTimeout(function() {
            try{
            apply_ace_system_settings()
            }catch(e){
              setTimeout(apply_ace_system_settings(),5000)//finally, wait for 5 more seconds
            }
          }, 5000);//then wait for 5 more seconds
    
        }
      }, 1000);  // first wait for 1 second
    }
  })
}

function configure_settings(){
  toggle_element('settings');
  if(!tag('settings').className.includes("hidden")){
    tag('ace-theme').focus();
    tag('settings-button').scrollIntoView(true);
    //console.log("fontSize",global_ace_options.fontSize)  
    tag("ace-font-size").value = global_ace_options.fontSize.replace("pt","")
    if(global_ace_options.wrap===false){
      tag("ace-word-wrap").value="no-wrap"
    }else if(global_ace_options.indentedSoftWrap){
      tag("ace-word-wrap").value="wrap-indented"
    }else{
      tag("ace-word-wrap").value="wrap"
    }
    //console.log("theme", global_ace_options.theme)
    //console.log("theme", global_ace_options.theme.split("/")[2])
    tag("ace-theme").value  = global_ace_options.theme.split("/")[2]
    tag("ace-line-numbers").checked = global_ace_options.showGutter
  }
}


function save_options(){
  global_ace_options.theme="ace/theme/" + tag("ace-theme").value
  global_ace_options.fontSize=tag("ace-font-size").value + "pt"
  global_ace_options.showGutter=tag("ace-line-numbers").checked
  switch(tag("ace-word-wrap").value){
    case "wrap":
      global_ace_options.wrap=true
      global_ace_options.indentedSoftWrap=false
      break
    case "wrap-indented":
      global_ace_options.wrap=true
      global_ace_options.indentedSoftWrap=true
      break
    default:  
      global_ace_options.wrap="off"
  }
  //console.log(global_ace_options)
  apply_editor_options(global_ace_options)
  const code = `function apply_ace_system_settings(){
    global_ace_options=${JSON.stringify(global_ace_options,null,4)}

    apply_editor_options(global_ace_options)
  }
  `
  save_module_to_workbook(code,"ACE System")
  hide_element('settings')
}

function apply_editor_options(options){
  for(const panel_name of code_panels){
    //console.log("updating options on ", panel_name)
    const editor = ace.edit(panel_name + "-content");
    editor.setOptions(options)
  }
}


async function submit_feedback(){
  // send feedback to the google form
  const message=[]
  if(tag("fb-type").value===""){message.push("<li>You must indicate the <b>type</b> of your feedback.</li>")}
  if(tag("fb-text").value===""){message.push("<li>You must provide the <b>text</b> of your feedback.</li>")}
  if(tag("fb-platform").value===""){message.push("<li>You must provide the <b>platform</b> you are using.</li>")}
  if(message.length===0){
    // ready to submit
    const form_data=[]
    form_data.push("entry.1033992853=")
    form_data.push(encodeURIComponent(tag("fb-type").value))
    form_data.push("&entry.482647522=")
    form_data.push(encodeURIComponent(tag("fb-platform").value))
    form_data.push("&entry.1230850697=")
    form_data.push(encodeURIComponent(tag("fb-email").value))
    form_data.push("&entry.1009690762=")
    form_data.push(encodeURIComponent(tag("fb-text").value))
    form_data.push("&entry.64124153=")
    form_data.push(encodeURIComponent("Addin"))

    tag("fb-message").innerHTML=""
    
    const options = {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: form_data.join("")
    };
    //console.log("form data",form_data.join(""))
    try{
    const response = await fetch("https://docs.google.com/forms/u/0/d/e/1FAIpQLSeRCZKHMU0HvXEiia8dn6hL0DfGoAH-PCJZolfiN1S_O90eSQ/formResponse", options)
    //console.log("reson",response.status)
    const data = await response.text()
    //console.log("data",data)
    tag("fb-message").style.color="green"

    //console.log("=====================",tag("fb-type").value)
    switch(tag("fb-type").value){
      case "Feature Resuest":
        message.push("Thanks for your feedback. Were not sure when we'll be updaing next but thanks for helping us understand your needs.")
        break
      case "Report Issue":
        message.push("Thanks for reporing this issue. While we cannot respond to every problem report, we'll do what we can.")
        break
      case "Offer to help with development":
        message.push("Thanks for offering to help on this project.  We'll take a look at your comment and get back to you--if you provided a valid email address.")
        break
      case "Praise for Addin":
        message.push("Thanks a million. We thrive on positive feedback!")
        break
      default:  
        message.push("Thanks for asking this question. While we cannot respond to every question, we'll do what we can.")
    }
    tag("fb-message").innerHTML=message.join("")
    tag("fb-message").scrollIntoView(true);
    setTimeout(function() {
      hide_element('survey')
      tag("fb-message").innerHTML=""
    }, 10000);
    }catch(e){
      //console.log("form error: ",e)
      tag("fb-message").style.color="red"
      tag("fb-message").innerHTML='Oops.  It looks as though there was a network error.  Your can tray again later or submit at our <a href="https://docs.google.com/forms/d/e/1FAIpQLSeRCZKHMU0HvXEiia8dn6hL0DfGoAH-PCJZolfiN1S_O90eSQ/viewform?usp=pp_url&entry.64124153=web" target="_blank">Google Form<a>.'  
      window.scrollTo(0,document.body.scrollHeight);
      tag("fb-message").scrollIntoView(true);
    }
    
  }else{
    tag("fb-message").style.color="red"
    tag("fb-message").innerHTML="<ul>" + message.join("") + "</ul>"
    tag("fb-message").scrollIntoView(true);
  }
  
}

function add_code_module(name,code){
  // a module built with whatever code is in default_code
  if(!code){// no code is pased in, determine which default code to import
    if(code_panels.length === 0){
      code=default_code()
    }else{  
      code=default_code("panel_" + name.toLowerCase().split(" ").join("_") + "_module")
    }
  }
  if(!name){name="code"}
  add_code_editor(name, code, "")
  hide_element("add-module")
  show_element("open-editor")
  show_panel(code_panels[code_panels.length-1])
  write_module_to_workbook(code,code_panels[code_panels.length-1])
}

async function get_code_from_gist(url){
  const response = await fetch(url.replace("gist.github.com", "gist.githubusercontent.com") + "/raw/?" + new Date())
  const data = await response.text()
  return data
}

async function import_code_module(url){
  const code = await get_code_from_gist(url)
  let name = null
  if(code.includes("ace.module:")){ // this is a function we can run directly and it as the comment
    try{
      name = JSON.parse(code.split("ace.module:")[1].split("*/")[0]).name
    }catch(e){
      //console.log("Error getting gist", e)
    }
  }
  if(!name){
    //either there was no comment to specify a name, or there was an error in reading it
    let x=1
    while(!!tag(panel_label_to_panel_name("Gist "+ x ))){
      x++
    }
    name="Gist " + x
  }
  add_code_module(name, code)
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
    //console.log("fetching")
    const response = await fetch(styles[style_name])
    const data = await response.text()
    //console.log("data",data)
    styles[style_name]=data
    if(integrate_now){
      document.getElementById("head_style").remove()
      document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + styles[style_name] + css_suffix + "</style>")
    }
  }
  
}

function set_style(style_name){

  //console.log("at set style", style_name)
  let css_sfx = css_suffix
  if(!style_name){
    style_name="system"
    css_sfx=""

  }

  if(styles[style_name].substr(0,8)==="https://"){
    // this style has not been fetched.  Get it now
    //console.log("in iff")
    get_style(style_name)
  }

  const style_tag = document.getElementById("head_style")
  if(style_tag.dataset.name!==style_name ){
    // only update the style tag if it is a differnt name
    style_tag.remove()
    document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="'+style_name+'">' + styles[style_name] + css_sfx + "</style>")
  }
  
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
  const panel_name="panel_listings"
 //console.log(1)
  if(!panels.includes){
    build_panel(panel_name)
  }
  

  // get the list of functions
  //###################################################### need to iterate over all modules
  const html=['<div onclick="close_canvas(\'' + panel_name + '\')" class="top-corner" style="padding:5px 5px 0 5px;margin:5px 15px 0 0; cursor:pointer"><i class="far fa-window-close fa-2x"></i></i></div><h2 style="margin:0 0 0 1rem">Active Automations</h2><ol>']
  //console.log("code panels", code_panels)
  for(const code_panel of code_panels){
    const code=tag(code_panel+"_div").firstChild.innerText
    const functions=get_function_names(code).function
    
    for(const func of functions){
      //console.log(func)
      let func_value = null
      if(code.includes("async function " + func + "(excel)")){
        //this is an async call to a funtion that interact with the workbook
        func_value=func + "(excel)"
        //console.log("adding",func)
      }else if(code.includes("function " + func + "()")){
        // this is a regular JS function with no params
        func_value=func+"()"
      }
      const function_text = window[func]+''
      //console.log(func_value, function_text.includes("ace.listing:"), function_text)
      if(func_value && function_text.includes("ace.listing:")){ // this is a function we can run directly and it as the comment
          //console.log("found a comment", func)
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
          ;console.log("ace.listing was not valid JSON", comment)        
        }
      }//for function on code page
    }//for code page 
  }

  if(html.length===1){
    //Did not find any properly configured fucntions
    html.push("There are currently no active automations in this workbook.")
  }else{
    html.push("</ul>")  
  }
  open_canvas("panel_listings",html.join(""))
}

function show_panel(panel_name){

  if(code_panels.includes(panel_name)){
    // set the size in case it is off
    tag(panel_name + "_editor-page").style.height = editor_height()
    try{
      ace.edit(panel_name + "-content").focus()
    }catch(e){
      ;console.log("could not access ace.  This is expected",e)
    }
  }

  if(panels.slice(0, 2).includes(panel_name) || code_panels.includes(panel_name)){
    set_style()
  }
  
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
      //console.log(" hiding", panel)
      tag(panel).style.display="none"  
    }
  }

  if(code_panels.includes(panel_name)){
    //focus the ace editor
    try{
      ace.edit(panel_name + "-content").focus()
    }catch(e){
      ;console.log("could not access ace.  This is expected",e)
    }
  }

}

function toggle_theme(panel_name){
  //console.log("changing theme")
  const editor = ace.edit(panel_name + "-content");
  let theme;

  switch (editor.getOptions().theme){
    case "ace/theme/tomorrow_night":
      theme = "ace/theme/tomorrow"
      break
    case "ace/theme/tomorrow":
      theme = "ace/theme/solarized_light"
      break
    default:
      theme = "ace/theme/tomorrow_night"
  }
  global_theme=theme
  editor.setOptions({
    theme: theme
  })
  
}

function toggle_wrap(panel_name){
  const editor = ace.edit(panel_name + "-content");
  if(editor.getOptions().wrap==="off"){
    editor.setOptions({wrap: true})
    globalThis=true
  }else{
    editor.setOptions({wrap: "off"})
    global_wrap="off"
  }
}

function add_code_editor(module_name, code, module_xmlid, settings, options_in){
  // settings are things gove is storing with the module
  // options are the options from the ace editor
  
  let options = global_ace_options


  // not currently handling options at the editor level, so this block is diabled
  // if(options_in){// default options for the editor
  //   options=options_in
  // }

console.log(module_name, "options", options)
  if(!settings){
    settings={cursorPosition:{row: 0, column: 0}}
  }
  
  //console.log("adding ace editor", module_name, module_xmlid)
  const panel_name = "panel_" + module_name.toLowerCase().split(" ").join("_") + "_module"
  code_panels.push(panel_name)
  panel_labels.push(panel_name_to_panel_label(panel_name))
  panels.push(panel_name)
  build_panel(panel_name)
  tag(panel_name).dataset.module_name = module_name
  tag(panel_name).dataset.module_xmlid = module_xmlid
  

 //console.log("initializing examples", tag(panel_name))
  
  tag(panel_name).appendChild(get_panel_selector(panel_name))
  const editor_container = document.createElement("div")
  editor_container.className=panel_name+"-content"
  

  let div = document.createElement("div");
  div.style.verticalAlign="middle"
  div.style.verticalAlign.padding=".2rem"
  div.innerHTML = '<button title="Save code to workbook" onclick="update_editor_script(\'' + panel_name + '\')">Save</button> <button  title="Save code to workbook and execute" onclick="code_runner(tag(\'' + panel_name + '_function-names' + '\').value,\'' + panel_name + '\')">Run</button> <select id="' + panel_name + '_function-names"></select>';
  div.style.height="22px"
  div.style.fontFamily="auto";
  div.style.fontSize = "1rem";
  div.style.padding=".2rem"
  div.style.backgroundColor = "DimGray";
  div.id=panel_name + "_editor-bar"
  editor_container.appendChild(div);


  //console.log("=======================================")
  //console.log(div);
  //console.log(div.clientHeight);

  const box = document.createElement("div");
  box.id = panel_name + "_editor-page";
  box.style.width = "100%";
  box.style.height = editor_height()
  box.style.display = "inline-block";
  box.style.position = "relative";

 //console.log("document",document.body.clientHeight);
 //console.log("scr",tag("panel_code_editor").Height);


  div = document.createElement("div");
  div.id = panel_name + "-content";
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
    const editor = ace.edit(panel_name + "-content");
    editor.getSession().setUseWorker(false);
    editor.on("blur", function () {
      window.event.target.parentElement.dataset.edited=true
    });
    editor.setTheme("ace/theme/solarized_light");
    if(!isNaN(options.fontSize)){
      options.fontSize += "pt"
    }
    editor.setOptions(options);
    editor.session.setMode("ace/mode/javascript");

    //console.log("settings", settings)
    editor.moveCursorTo(settings.cursorPosition.row, settings.cursorPosition.column)

    editor.commands.addCommand({  // toggle word wrap
      name: "wrap",
      bindKey: {win: "Ctrl-shift-w", mac: "Command-shift-w"},
      exec: function(editor) {
        for(const panel of code_panels){
          if(tag(panel).style.display==="block"){
            // we found the one that is visible
            toggle_wrap(panel)
            break//exit the loop
          }
        }
        
      }
    })

    editor.commands.addCommand({  // toggle Theme
      name: "theme",
      bindKey: {win: "Ctrl-shift-t", mac: "Command-shift-t"},
      exec: function(editor) {
        for(const panel of code_panels){
          if(tag(panel).style.display==="block"){
            // we found the one that is visible
            toggle_theme(panel)
            break//exit the loop
          }
        }
        
      }
    })

    editor.commands.addCommand({  // can't make ctrl+shift work.  I expect that the addin environment is trappping that keystroke
      name: "save",
      bindKey: {win: "Ctrl-shift+s", mac: "Command-shift-s"},
      exec: function(editor) {
        for(const panel of code_panels){
          if(tag(panel).style.display==="block"){
            // we found the one that is visible
            update_editor_script(panel)
            break//exit the loop
          }
        }
        
      }
    })
    editor.commands.addCommand({  // could do ctrl+r but want to be parallel with save
      name: "run",
      bindKey: {win: "Ctrl-e", mac: "Command-e"},
      exec: function(editor) {
        for(const panel of code_panels){
          if(tag(panel).style.display==="block"){
            // we found the one that is visible
            code_runner(tag(panel + '_function-names').value, panel)
            break//exit the loop
          }
        }
      }
    })

    editor.commands.addCommand({  // could do ctrl+r but want to be parallel with save
      name: "run_shift",
      bindKey: {win: "Ctrl-shift-e", mac: "Command-shift-e"},
      exec: function(editor) {
        for(const panel of code_panels){
          if(tag(panel).style.display==="block"){
            // we found the one that is visible
            code_runner(tag(panel + '_function-names').value, panel)
            break//exit the loop
          }
        }
      }
    })

  });

  

  editor_container.style.display = "block";

  div = document.createElement("div");
  div.id = panel_name + "_div";      
  const script = document.createElement("script");
  script.innerHTML = code;
  div.appendChild(script);
  editor_container.appendChild(div);


  tag(panel_name).appendChild(editor_container)

  load_function_names_select(code, panel_name)
  if(settings.func){
    tag(panel_name + "_function-names").value=settings.func

  }

  // AutoExecutable function
 //console.log("about to autoexec")
  try{
   //console.log("in try")
    auto_exec()
  }catch(e){
   ;console.log("catch",e)
  }  
}

function code_runner(script_name,panel_name){
  //console.log(script_name,panel_name)
  if (tag(panel_name + "-content").dataset.edited==="true"){
    update_editor_script(panel_name)
  }
  if(script_name.includes("(excel)")){
    setTimeout("Excel.run(" + script_name.split("(")[0] + ")", 0) //run the function
  }else{
    setTimeout(script_name, 0) //run the function
  }
}

function init_examples(){
  const panel_name="panel_examples"
  build_panel(panel_name)
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
  // these examples should be made in script lab to have the right format
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
          const editor = ace.edit("editor" + id);
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
  // this is the one used by the examples page
  //console.log("script" + id);
  script_div = document.getElementById("script" + id);
  script_div.innerHTML = "";
  const editor = ace.edit("editor" + id);
  const script = document.createElement("script");
  script.innerHTML = editor.getValue();
  script_div.appendChild(script);
}


function update_editor_script(panel_name) {
  // read the script for an ace editor and write it to the DOM
  // also saves the module to the custom properties
  //console.log("at update_editor_script", panel_name)
  // set the size of the editor in case there was a prior zoom
  tag(panel_name + "_editor-page").style.height = editor_height()

  script_div = document.getElementById(panel_name + "_div");
  script_div.innerHTML = "";

  // update the script_div for one code panel
  const code = ace.edit(panel_name + "-content").getValue();
  //console.log(code)
  const script = document.createElement("script");
  script.innerHTML = code
  script_div.appendChild(script);
  load_function_names_select(code, panel_name)
  write_module_to_workbook(code, panel_name)
  
}
function write_module_to_workbook(code, panel_name){
  let options = {theme:global_theme,wrap:global_wrap,fontSize:"14pt"}
  const settings = {
    cursorPosition:{row: 0, column: 0},
    func:tag(panel_name + "_function-names").value
  }

  try{
    const editor = ace.edit(panel_name + "-content") 
    settings.cursorPosition = editor.getCursorPosition()
    options=editor.getOptions()
  }catch(e){
    //console.log("This is an expected error: ace editor not yet built",e)
  }

  //console.log("writing options", JSON.stringify(options))
  
  const module_name = tag(panel_name).dataset.module_name
  let xmlid = tag(panel_name).dataset.module_xmlid

  if (xmlid){  // workbook as already been saved and has and xmlid
    //console.log("saving an existing book")
    save_module_to_workbook(code, module_name, settings, xmlid, options)
  }else{  // workbook not yet saved, function call will return an xmlid
    xmlid=save_module_to_workbook(code, module_name, settings, xmlid, options)
    //console.log("xmlid", xmlid)
    tag(panel_name).dataset.module_xmlid = xmlid
  }
}

function save_module_to_workbook(code, module_name, settings, xmlid, options){
  // options are currently being ignored, they are handled globally instead of at the module level
  // save to workbook
  Excel.run(async (excel)=>{
    //console.log("saving code", panel_name) 
    if(!settings){
      settings = {
        cursorPosition:{row: 0, column: 0},
      }
    }
    //The next line has been disabled because we are not currently maintaining options at the module level
    //const module_xml = "<module xmlns='http://schemas.gove.net/code/1.0'><name>"+module_name+"</name><settings>"+btoa(JSON.stringify(settings))+"</settings><options>"+btoa(JSON.stringify(options))+"</options><code>"+btoa(code)+"</code></module>"

    // save module without options
    const module_xml = "<module xmlns='http://schemas.gove.net/code/1.0'><name>"+module_name+"</name><settings>"+btoa(JSON.stringify(settings))+"</settings><options>"+btoa(null)+"</options><code>"+btoa(code)+"</code></module>"
    if(xmlid){
      //console.log("updating xml", xmlid)
      const customXmlPart = excel.workbook.customXmlParts.getItem(xmlid);
      customXmlPart.setXml(module_xml)
      excel.sync()
      return null
    }else{
      //console.log("creating xml")
      const customXmlPart = excel.workbook.customXmlParts.add(module_xml);
      customXmlPart.load("id");
      await excel.sync();

      //console.log("customXmlPart",customXmlPart.getXml())
      // this is a newly created module and needs to have a custom xmlid part made for it
      code_module_ids.push(customXmlPart.id)                   // add the id to the list of ids
      const settings = excel.workbook.settings;
      settings.add("code_module_ids", code_module_ids);  // adds or sets the value
      await excel.sync()
      return customXmlPart.id
  
    }
    
  })

}



function load_function_names_select(code,panel_name){// reads the function names from the code and puts them in the function name select
  const selectElement=tag(panel_name + "_function-names")
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
  const panel_label=panel_name_to_panel_label(panel)
  const sel = document.createElement("select")
  //console.log("appending panel=====", panel)
  if(!panel_labels.includes(panel_label)){
    panel_labels.push(panel_label)
  }
  sel.className="panel-selector"
  // put the options in this panel selector
  update_panel_selector(sel)

  // update all others panel selectirs
  for(const selector of document.getElementsByClassName("panel-selector")){
    update_panel_selector(selector)
  }
  
  sel.value=panel
  sel.style.height="40px"
  sel.id="selector_"+panel
  sel.onchange = select_page
  return sel
}

function update_panel_selector(sel){
  // put the proper choices in a panel selector
  while(sel.length>0){
     sel.remove(0)
  }
  for (let i=0; i<panel_labels.length; i++) {
    var option = document.createElement("option");
    option.value = panel_label_to_panel_name(panel_labels[i]) 
   //console.log("-->", option.value)
    option.text = panel_labels[i];
    option.className="panel-selector-option"
    sel.appendChild(option);
  }
}



function panel_label_to_panel_name(panel_label){
  return "panel_" + panel_label.toLocaleLowerCase().split(" ").join("_");
}

function panel_name_to_panel_label(panel_name){
  let panel_label = panel_name.replace("panel_","")
  panel_label=panel_label.split("_").join(" ") 
  return titleCase(panel_label)
}

function titleCase(str) {
  str = str.toLowerCase().split(' ');
  for (var i = 0; i < str.length; i++) {
    str[i] = str[i].charAt(0).toUpperCase() + str[i].slice(1); 
  }
  return str.join(' ');
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









function show_element(tag_id){
  // removes the hidden class from a tag's css
  tag(tag_id).className=tag(tag_id).className.replaceAll("hidden","")
}

function hide_element(tag_id){
  // adds the hidden class from a tag's css
  //console.log("tag_id",tag_id)
  //console.log("tag(tag_id)",tag(tag_id))
  //console.log("tag(tag_id).className",tag(tag_id).className)
  if(tag(tag_id).className){
    if(!tag(tag_id).className.includes("hidden")){
      tag(tag_id).className=(tag(tag_id).className + " hidden").trim()
    }
  }else{
    tag(tag_id).className="hidden"
  }
  
}

function toggle_element(tag_id){
  // adds the hidden class from a tag's css
  if(tag(tag_id).className.includes("hidden")){
    show_element(tag_id)
  }else{
    hide_element(tag_id)
  }
}

function editor_height(){
  return (window.innerHeight-73)+"px"
}

function default_code(panel_name){
  let code=`async function write_timestamp(excel){
    /*ace.listing:{"name":"Timestamp","description":"This sample function records the current time in the selected cells"}*/
  excel.workbook.getSelectedRange().values = new Date();
  await excel.sync();
}

function auto_exec(){
  // This function is called when the addin opens.
  // un-comment a line below to take action on open.

  // open_automations() // displays a list of functions for a user
`
  if(panel_name){
    code += `  // show_panel('${panel_name}')      // shows this code editor
}`
  }else{
    code += `  // open_editor()      // shows the code editor
}`
  }
  return code
}