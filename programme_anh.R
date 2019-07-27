# install.packages("xlsx")
library("xlsx")
library(gdata)
library(XLConnect)

chemin = "/home/quyen/Documents/_projets/rlang/Anh/"

#----------------------------------------------------------------------
# Declaration des fonctions
#----------------------------------------------------------------------
xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows, colIndex=1)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}

#----------------------------------------------------------------------
# Declaration des parametres
#----------------------------------------------------------------------
liste_titre_full = c("Alarmes avérées",
                     "Alarmes non avérées comportementales",
                     "alarme qualifiee")
  
fichier_input = paste0(chemin,"data.xls")
fichier_out = paste0(chemin,"data_out.xlsx")

Tab_data = as.matrix(read.xls(fichier_input,sheet=1))

liste_onglets = unique(Tab_data[,"DEPARTEMENT"])
Nb_onglets = length(liste_onglets)

wb = createWorkbook(type="xlsx")

for(i in 1:Nb_onglets){
  onglet = liste_onglets[i]
  ind_onglet = which(Tab_data[,"DEPARTEMENT"] == onglet)
  
  liste_titre = unique(Tab_data[ind_onglet,"TYPE"])
  Nb_titre = length(liste_titre)
  
  TITLE_STYLE <- CellStyle(wb)+ Font(wb,  heightInPoints=16,  color="blue", isBold=TRUE, underline=1)
  
  sheet <- createSheet(wb, sheetName = onglet)
  
  index_row = 1
  for(t in 1:Nb_titre){
    titre = liste_titre[t]
    xlsx.addTitle(sheet, rowIndex=index_row, title=titre,
                  titleStyle = TITLE_STYLE)
    
    ind_titre = which((Tab_data[,"TYPE"] == titre)&(Tab_data[,"DEPARTEMENT"] == onglet))
    Nb_elements = length(ind_titre)
    
    index_row = index_row+1
    
    index_col = 1
    NOABREGE = Tab_data[ind_titre,"NOABREGE"]
    addDataFrame(NOABREGE, sheet=sheet, startRow=index_row, startCol=index_col, row.names = F, col.names = F)
    
    index_col = index_col+1
    NOM = Tab_data[ind_titre,"NOM"]
    addDataFrame(NOM, sheet=sheet, startRow=index_row, startCol=index_col, row.names = F, col.names = F)
    
    index_col = index_col+1
    DATE = Tab_data[ind_titre,"DATE"]
    addDataFrame(DATE, sheet=sheet, startRow=index_row, startCol=index_col, row.names = F, col.names = F)
    
    index_col = index_col+1
    CODE = Tab_data[ind_titre,"CODE"]
    addDataFrame(CODE, sheet=sheet, startRow=index_row, startCol=index_col, row.names = F, col.names = F)
    
    index_col = index_col+1
    LIBELLE = Tab_data[ind_titre,"LIBELLE"]
    addDataFrame(LIBELLE, sheet=sheet, startRow=index_row, startCol=index_col, row.names = F, col.names = F)
    
    tmp = length(NOABREGE)
    
    index_row = index_row + tmp + 1
  }
  
  
  saveWorkbook(wb, fichier_out)
  
}



