const { app, BrowserWindow,Menu,ipcMain } = require('electron');
const path = require('path');

//Windows vars
let loginWindow
let homeWindow
let uploadExcelWindow

/* Create windows */
const createLoginWindow = () =>{
  loginWindow = new BrowserWindow({
    width:600,
    height:600,
    minWidth:600,
    minHeight:600,
    resizable:false,
    title:'RR.HH APP',
    icon:__dirname+'./icon.ico',
    autoHideMenuBar:true,
    webPreferences: {
        nodeIntegration: true,
        contextIsolation: false,
        enableRemoteModule: true,
         
    }
  })
  loginWindow.on('closed',()=>{app.quit()});
  loginWindow.loadFile('./src/login.html');
}
const createHomeWindow = ()=>{
  homeWindow = new BrowserWindow({
    width:600,
    height:600,
    minWidth:600,
    minHeight:600,
    resizable:false,
    title:'RR.HH APP',
    icon:__dirname+'./icon.ico',
    autoHideMenuBar:true,
    webPreferences: {
        nodeIntegration: true,
        contextIsolation: false,
        enableRemoteModule: true,
         
    }
  })
  homeWindow.on('closed',()=>{app.quit()});
  homeWindow.loadFile('./src/index.html');
}
const createUploadExcelWindow = () => {
    uploadExcelWindow = new BrowserWindow({
      width: 1600,
      height: 900,
      minWidth: 1280,
      minHeight: 720,
      title: 'RR.HH APP',
      icon:__dirname+'./icon.ico',
      webPreferences:{
        nodeIntegration:true,
        contextIsolation: false,
        enableRemoteModule:true,
      }
     
    })
    
    /* const homeMenu = Menu.buildFromTemplate(templateMenu)
    Menu.setApplicationMenu(homeMenu) */
    uploadExcelWindow.loadFile('./src/uploadExcel.html');

}

const createUploadedFilesWindow = () =>{
  const uploadedFilesWindow = new BrowserWindow({
    width:1280,
    height:720,
    resizable:false,
    title:'RR.HH APP',
    icon:__dirname+'./icon.ico',
    autoHideMenuBar:true
  })

  uploadedFilesWindow.loadFile('./src/subidos.html')
}

const createHelpWindow = ()=>{
  const helpWindow = new BrowserWindow({
    width:800,
    height:600,
    resizable:false,
    title: 'Ayuda',
    icon:__dirname+'./icon.ico',
    autoHideMenuBar:true
  })
}



/* Menus */
/* Menu superior del home Windows */
const templateMenu = [
  {
    label:'Subidos',
    click(){
      createUploadedFilesWindow();
    }
  },
  {
    label:'Ayuda',
    click(){
      createHelpWindow();
    }
  },

]



//Eventos
app.on('ready',()=>{
  createLoginWindow();
})
ipcMain.on("login",()=>{
  createHomeWindow();
  loginWindow.hide();
})

ipcMain.on("open:buscarBoletasWindow",()=>{
  createUploadedFilesWindow();
})

ipcMain.on("open:emitirBoletasWindow",()=>{
  createUploadExcelWindow();
})
