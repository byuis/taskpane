//  these are the functions that are documented for use by users of the addin
const global={}
const default_style_name="marx"

async function load_from_gist(gist_url){
    //takes a gist URL and loads it's content in a script window
    const code=await get_code_from_gist(gist_url)
    console.log(code)
    const script = document.createElement("script")
    script.innerHTML = code
    script.id="gist-import-script"
    document.body.appendChild(script)
    tag("gist-import-script").remove()
  
}


function set_css(user_css){
    css_suffix=user_css
}

function add_library(url){
    // adds a JS library to the head section of the HTML sheet
    const library = document.createElement('script');
    library.setAttribute('src',url);
    console.log("library",library)
    document.head.appendChild(library);
}

function set_theme(theme_name){
    set_style(theme_name)
}

function list_themes(){
    for(const [theme, url] of Object.entries(styles)){
        console.log(theme, url)
    }
}


function tag(id){
    // a short way to get an element by ID
    return document.getElementById(id)
}

function close_canvas(){
    panel_stack.pop()
    show_panel(panel_stack.pop())
}

function open_editor(){
    show_panel(code_panels[0])
}

function open_automations(){
    show_automations()
}

function reset(){
    show_panel("panel_introduction")
}

function open_canvas(panel, proc_or_html, style_name){
    if(style_name){
        set_style(style_name)
    }

    if(!tag(panel)){
        build_panel(panel)
    }

    if(!panels.includes(panel)){
        panels.push(panel)
    }

    show_panel(panel)

    if(typeof proc_or_html === "function"){
        proc_or_html()
    }else{
        if(proc_or_html){
            tag(panel).innerHTML=proc_or_html
        }
    }
}
  
function show_examples(){
    const panel_name="panel_examples"
    set_style()
    show_panel(panel_name)
  
  }