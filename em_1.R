rm(list=ls())

if(is.installed("xlsx")){
  install.packages("xlsx")
}

if(is.installed(XLConnect)){
  install.packages(XLConnect)
}

if(is.installed(gdata)){
  install.packages(gdata)
}

library(XLConnect)
library(gdata)
library("xlsx")


chemin = "/home/lpc/"

#----------------------------------------------------------------------
# Declaration des fonctions
#----------------------------------------------------------------------
xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}

#----------------------------------------------------------------------
# Declaration des parametres
#----------------------------------------------------------------------
colIndex = 1:5
liste_titre_full = c("Alarmes averees",
                     "Alarmes non averees comportementales",
                     "Alarme non qualifiee",
                     "Alarme techniques")
Nb_titre_full = length(liste_titre_full)

fichier_input = paste0(chemin,"data.xls")
fichier_out = paste0(chemin,"data_out_1.xlsx")

Tab_data = as.matrix(read.xls(fichier_input,sheet=1))

liste_onglets = unique(Tab_data[,"DEPARTEMENT"])
Nb_onglets = length(liste_onglets)

wb = createWorkbook(type="xlsx")
Nb_onglets = 2
for(i in 1:Nb_onglets){
  onglet = liste_onglets[i]
  ind_onglet = which(Tab_data[,"DEPARTEMENT"] == onglet)
  
  liste_titre = unique(Tab_data[ind_onglet,"TYPE"])
  Nb_titre = length(liste_titre)
  
  # Fill(foregroundColor  = "blue", backgroundColor = "blue",pattern="SOLID_FOREGROUND") + 
  TITLE_STYLE <- CellStyle(wb) + Font(wb,  heightInPoints=12,  color="blue", isBold=TRUE, name = "Calibri", underline = NULL)
  
  sheet <- createSheet(wb, sheetName = onglet)
  
  index_row = 1
  for(t in 1:Nb_titre_full){
    titre = liste_titre_full[t]
    xlsx.addTitle(sheet, rowIndex=index_row, title=titre,
                  titleStyle = TITLE_STYLE)
    # mergeCells(wb, i, reference = paste0("A",index_row,":E",5))
    
    ind_titre = which((Tab_data[,"TYPE"] == titre)&(Tab_data[,"DEPARTEMENT"] == onglet))
    Nb_elements = length(ind_titre)
    
    index_row = index_row+1
    if(Nb_elements == 0){
      index_col = 1
      addDataFrame("Pas de données", sheet=sheet, startRow=index_row, startCol=index_col, row.names = F, col.names = F)
      index_row = index_row + 1
    }else{
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
      
      index_row = index_row + tmp 
    }
    
    
  }
  
  
  saveWorkbook(wb, fichier_out)
  
}