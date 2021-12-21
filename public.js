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

function open_output(){
    show_panel("panel_output")
}


function open_automations(){
    show_automations()
}

function reset(){
    show_panel("panel_introduction")
}

function show_html(html){
    //A simple function that is mapped differntly for examples than for modules
    //this is the module mapping
    open_canvas("html", html)
}

function open_canvas(panel_name, html, style_name){
    if(style_name){
        set_style(style_name)
    }

    if(!tag(panel_name)){
        build_panel(panel_name)
    }

    if(!panels.includes(panel_name)){
        panels.push(panel_name)
    }

    show_panel(panel_name)

    if(html){
        tag(panel_name).innerHTML=panel_close_button(panel_name) + html
    }
}
function print(data, heading){
    //if(!header && )
    if(!tag("panel_output").lastChild.lastChild.firstChild.tagName && !heading){
        //no output here, need a headdng
        heading=""
    }
    if(heading!==undefined){
        // there is a header, so make a new block
        console.log("at data")
        const div = document.createElement("div")
        div.className="ace-output"
        const header = document.createElement("div")
        header.className="ace-output-header"  
        const d = new Date()
        let ampm=" am"
        let hours=d.getHours()
        if(hours >11){
            ampm="pm"
            if(hours>12){
                hours=hours-12
            }
        }
        header.innerHTML = '<span class="ace-output-time">' + hours + ":" + ("0"+d.getMinutes()).slice(-2) + ":" + ("0"+d.getSeconds()).slice(-2) + ampm + "</span> " + heading + '<div class="ace-output-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>'
        const body = document.createElement("div")
        body.className="ace-output-body"  
        body.innerHTML = '<div style="margin:0;font-family: monospace;">' + data.replaceAll("\n","<br />")  + "<br />"+ "</div>"
        div.appendChild(header)
        div.appendChild(body)
        tag("panel_output").appendChild(div)
    }else{
        // no header provided, append to most recently added
        tag("panel_output").lastChild.lastChild.firstChild.innerHTML += data.replaceAll("\n","<br />") + "<br />"
    }

  
}
  
function alert(data, heading){
    if(tag("ace-alert")){tag("ace-alert").remove()}
    if(!heading){heading="System Message"}
    const div = document.createElement("div")
    div.className="ace-alert"
    div.id='ace-alert'
    const header = document.createElement("div")
    header.className="ace-alert-header"  
    header.innerHTML = heading + '<div class="ace-alert-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>'
    const body = document.createElement("div")
    body.className="ace-alert-body"  
    body.innerHTML = data
    div.appendChild(header)
    div.appendChild(body)
    document.body.appendChild(div)

}

function show_examples(){
    const panel_name="panel_examples"
    set_style()
    show_panel(panel_name)
  
  }