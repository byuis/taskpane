//  these are the functions that are documented for use by users of the addin
const global={}
const default_style_name="marx"

function set_css(user_css){
    css_suffix=user_css
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
    show_panel("panel_introduction")
}

function open_editor(){
    show_panel("panel_code_editor")
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
  