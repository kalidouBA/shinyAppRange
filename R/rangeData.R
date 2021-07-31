require(shiny)
require(shinyFiles)
require(shinydashboard)
require(gtools)
require(xlsx)
require(WriteXLS)
require(openxlsx)
require(rsconnect)
require(progress)
require(shinybusy)
require(shinyalert)
require(devtools)
require(roxygen2)

rangeData = function(numCols = 3) {

  volumes = c('R Installation'=R.home())
  viewer = shiny::dialogViewer("test", width = 1200, height = 1100)
  shiny::runGadget(shiny::shinyApp(ui = dashboardPage(
    dashboardHeader(title = "Rangement"),
    dashboardSidebar(),
    dashboardBody(
      fluidRow(
        useShinyalert(),
        box(
          height = 310,
          title = "Dossier de travail", status = "primary", solidHeader = TRUE,
          "Chargement du données contenant les données", br(),
          hr(),
          shinyDirButton("dir", label = 'Choisir le dossier contenant les données', "Upload",
                         style="color: #fff; background-color: #6495ED; border-color: #c34113;
                             border-radius: 10px;
                             border-width: 2px"),
          hr(),
          verbatimTextOutput("dir", placeholder = TRUE)
        ),
        box(
          height = 310,
          title = "Sauvegarde", status = "primary", solidHeader = TRUE,
          "Création du fichier de sauvegarde", br(), div(HTML("<em>Example name file: file1 or file.1 or file-1 or file_1</em>")),
          hr(),
          textInput("file_name","Non du ficher de sauvegarde"),
          actionButton("save","Sauvegarde"),
          hr()
        )
      )
    )
  ),
  server = function(input, output) {
    start.time = Sys.time()
    observeEvent(input$save,{
      SaveRData()
    })
    shinyDirChoose(
      input,
      'dir',
      roots = c(home = '/'),
      filetypes = c('', 'txt', 'bigWig', "tsv", "csv", "bw")
    )

    global = reactiveValues(datapath = getwd())

    dir = reactive(input$dir)

    output$dir = renderText({
      global$datapath
    })

    observeEvent(ignoreNULL = TRUE,
                 eventExpr = {
                   input$dir
                 },
                 handlerExpr = {
                   if (!"path" %in% names(dir())) return()
                   home = normalizePath("/")
                   global$datapath =
                     file.path(home, paste(unlist(dir()$path[-1]), collapse = .Platform$file.sep))
                 })


    SaveRData = reactive({
      if(!(is.null(input$save) && is.null(input$file_name))){
        file_name = paste0(global$datapath,"/",input$file_name,".xlsx")

        split.path = strsplit(global$datapath,"/")[[1]]
        directory.name = tail(split.path,1)
        length.path = length(split.path)

        # On réccupère tous les dossiers des contenus
        CDD = list.dirs(path = global$datapath)
        CDD = mixedsort(sort(CDD))

        wb = openxlsx::createWorkbook(paste0(global$datapath,"/Excel_toutes_métriques_tous_les_essais_v2_",directory.name,".xlsx"))
        sheet.TUG = addWorksheet(wb, "20_TUG")
        startCol.20_TUG = 1

        sheet.Walk_metrics_20_TUG = addWorksheet(wb, "Walk_metrics_20_TUG")
        startCol.metrics_20_TUG = 1

        sheet.SWAY_Baseline = addWorksheet(wb, "SWAY_Baseline")
        startCol.SWAY_Baseline = 1

        sheet.SWAY_Post_TUG = addWorksheet(wb, "SWAY_Post_TUG")
        startCol.SWAY_Post_TUG = 1

        sheet.10MPS_Baseline = addWorksheet(wb, "10MPS_Baseline")
        startCol.10MPS_Baseline = 1

        sheet.10MVmax_Baseline = addWorksheet(wb, "10MVmax_Baseline")
        startCol.10MVmax_Baseline = 1

        sheet.10MPS_Post_TUG = addWorksheet(wb, "10MPS_Post_TUG")
        startCol.10MPS_Post_TUG = 1

        sheet.10MVmax_Post_TUG = addWorksheet(wb, "10MVmax_Post_TUG")
        startCol.10MVmax_Post_TUG = 1


        n_iter = length(CDD)

        # Initializes the progress bar
        pb = progress_bar$new(format = "(:spin) [:bar] :percent [Elapsed time: :elapsedfull || Estimated time remaining: :eta]",
                               total = n_iter,
                               complete = "=",   # Completion bar character
                               incomplete = "-", # Incomplete bar character
                               current = ">",    # Current bar character
                               clear = FALSE,    # If TRUE, clears the bar when finish
                               width = 100)
        k = 1
        numFile = NA
        beginCol = NA
        data.in.file = NA
        for (elt in CDD) {
          show_modal_spinner()

          ## TUG
          if (identical(toupper(strsplit(elt,split = "/")[[1]][length.path+1]), "TUG")
              & !is.na(strsplit(elt,split = "/")[[1]][length.path+2])
              & length(strsplit(elt,split = "/")[[1]]) == length.path+2) {
            if ( identical(strsplit(strsplit(elt,split = "/")[[1]][length.path+2]," ")[[1]][1], "ESSAI")) {
              pathDir.files = list.files(path = elt,pattern="*_Trial.csv")
              data.in.file = read.table(paste0(elt,"/",pathDir.files),header=FALSE,
                                        sep = ";",quote = "\"",
                                        na.strings =" ", stringsAsFactors= F,
                                        col.names = paste0("V",seq_len(50)),fill = TRUE)
              data.in.file = data.frame(data.in.file)
              data.in.file = data.in.file[,colSums(is.na(data.in.file))<nrow(data.in.file)]
              numFile = 1
              startCol.20_TUG = startCol.20_TUG + ncol(data.in.file) + 2
              beginCol = startCol.20_TUG
            }
            else{
              dir.dirs = list.dirs(path = elt)
              dir.dirs = mixedsort(sort(dir.dirs))
              for(dir in dir.dirs){
                if (length(strsplit(dir,split = "/")[[1]]) == length.path+3) {
                  pathDir.files = list.files(path = dir,pattern="*_Trial.csv")
                  data.in.file = read.table(paste0(dir,"/",pathDir.files),header=FALSE,
                                            sep = ";",quote = "\"",
                                            na.strings =" ", stringsAsFactors= F,
                                            col.names = paste0("V",seq_len(50)),fill = TRUE)
                  data.in.file = data.frame(data.in.file)
                  data.in.file = data.in.file[,colSums(is.na(data.in.file))<nrow(data.in.file)]
                  numFile = 2
                  startCol.metrics_20_TUG = startCol.metrics_20_TUG + ncol(data.in.file) + 2
                  beginCol = startCol.metrics_20_TUG
                }
              }
            }


            ## SWAY
          }else if(identical(strsplit(elt,split = "/")[[1]][length.path+1], "SWAY")
                   & !is.na(strsplit(elt,split = "/")[[1]][length.path+2])
                   & length(strsplit(elt,split = "/")[[1]])>length.path+3) {
            pathDir = paste0(global$datapath,"/",
                             paste(strsplit(elt,split = "/")[[1]][(length.path+1):(length.path+3)],collapse = "/"))

            pathDir.files = list.files(path = pathDir,pattern="*_Trial.csv")
            data.in.file = read.table(paste0(pathDir,"/",pathDir.files),header=FALSE,
                                      sep = ";",quote = "\"",
                                      na.strings =" ", stringsAsFactors= F,
                                      col.names = paste0("V",seq_len(50)),fill = TRUE)
            data.in.file = data.frame(data.in.file)
            data.in.file = data.in.file[,colSums(is.na(data.in.file))<nrow(data.in.file)]
            ifelse (identical(strsplit(elt,split = "/")[[1]][length.path+2], "Baseline"),
                    { numFile = 3
                      startCol.SWAY_Baseline = startCol.SWAY_Baseline + ncol(data.in.file) + 2
                      beginCol = startCol.SWAY_Baseline
                    },
                    {
                      numFile = 4
                      startCol.SWAY_Post_TUG = startCol.SWAY_Post_TUG + ncol(data.in.file) + 2
                      beginCol = startCol.SWAY_Post_TUG
                      }
            )
          }

          ##  WALK
          else if(identical(strsplit(elt,split = "/")[[1]][length.path+1], "WALK")
                  & !is.na(strsplit(elt,split = "/")[[1]][length.path+3])
                  & length(strsplit(elt,split = "/")[[1]])>length.path+4) {
            pathDir = paste0(global$datapath,"/",paste(strsplit(elt,split = "/")[[1]][(length.path+1):(length.path+4)],collapse = "/"))
            pathDir.files = list.files(path = pathDir,pattern="*_Trial.csv")

            # PS
            if(identical(strsplit(pathDir,"/")[[1]][length.path+2],"PS")){
              data.in.file = read.table(paste0(pathDir,"/",pathDir.files),header=FALSE,
                                        sep = ";",quote = "\"",
                                        na.strings =" ", stringsAsFactors= F,
                                        col.names = paste0("V",seq_len(50)),fill = TRUE)
              data.in.file = data.frame(data.in.file)
              data.in.file = data.in.file[,colSums(is.na(data.in.file))<nrow(data.in.file)]
              ifelse(identical(strsplit(elt,split = "/")[[1]][length.path+3], "Baseline"),
                     {
                       numFile = 5
                       startCol.10MPS_Baseline = startCol.10MPS_Baseline + ncol(data.in.file) + 2
                       beginCol = startCol.10MPS_Baseline
                     },
                     {
                       numFile = 7
                       startCol.10MPS_Post_TUG = startCol.10MPS_Post_TUG + ncol(data.in.file) + 2
                       beginCol = startCol.10MPS_Post_TUG
                       })
            }

            # Vmax
            else if(identical(strsplit(pathDir,"/")[[1]][length.path+2],"Vmax")){
              if(identical(strsplit(elt,split = "/")[[1]][length.path+3], "Baseline")){

                data.in.file = read.table(paste0(pathDir,"/",pathDir.files),header=FALSE,
                                          sep = ";",quote = "\"",
                                          na.strings =" ", stringsAsFactors= F,
                                          col.names = paste0("V",seq_len(50)),fill = TRUE)
                data.in.file = data.frame(data.in.file)
                data.in.file = data.in.file[,colSums(is.na(data.in.file))<nrow(data.in.file)]
                numFile = 6
                startCol.10MVmax_Baseline = startCol.10MVmax_Baseline + ncol(data.in.file) + 2
                beginCol = startCol.10MVmax_Baseline
              }
              else{
                data.in.file = read.table(paste0(pathDir,"/",pathDir.files),header=FALSE,
                                          sep = ";",quote = "\"",
                                          na.strings =" ", stringsAsFactors= F,
                                          col.names = paste0("V",seq_len(50)),fill = TRUE)
                data.in.file = data.frame(data.in.file)
                data.in.file = data.in.file[,colSums(is.na(data.in.file))<nrow(data.in.file)]
              }
              numFile = 8
              startCol.10MVmax_Post_TUG = startCol.10MVmax_Post_TUG + ncol(data.in.file) + 2
              beginCol = startCol.10MVmax_Post_TUG
            }
          }
          print(dim(data.in.file))
          writeData(wb,numFile, data.in.file, colNames = FALSE , startCol = beginCol)

          # Sets the progress bar to the current state
          pb$tick()
          end.time = Sys.time()
          time.taken = end.time - start.time
          print(time.taken)
          pctg = paste(round(k/n_iter *100, 0), "% completed")
          k = k+1
        }
        openxlsx::saveWorkbook(wb, file = file_name,overwrite = TRUE)
        if(k == length(CDD)+1)shinyalert("Sauvegarde réussie!", "Le fichier excel est situé dans le dossier de données", type = "success",imageUrl = "https://jeroen.github.io/images/banana.gif",
                                         imageHeight = 70,imageWidth = 70)
        else shinyalert("Vérifier le contenu du fichier", "Le fichier excel est situé dans le dossier de données", type = "warning")
        remove_modal_spinner()
      }
      if(k == 1)
        shinyalert("Sauvegarde échouée!", "Le fichier excel n'a pas pu etre créé", type = "error")
    })

  }), viewer = viewer,
  stopOnCancel = FALSE)
}
