#read in libraries
library(dplyr)
library(openxlsx)
library(reshape2)
library(ggplot2)

#read in files
file_path = "C:/Users/Penny/Documents/DSHub/HubRoster/Data Science Roster_ALL.xlsx"
data <- read.xlsx(file_path, sheet = "Responses")

method_experts<-which(data=="Expert", arr.ind=TRUE)
appli_res<-which(data=="Developer/Researcher", arr.ind=TRUE)

#organize data for roster on website
output_strings<- list()
for (i in 1:nrow(data)){
  name <- paste(data$FirstName[i],data$LastName[i])
  if (is.na(data$Photo_Id[i])){
    image_string<-c("<i class='fa-regular fa-user fa-9x'></i>")
  }
  else{
    image_string<-paste("![",name,"](images/",name,".jpg){.headshot}",sep="")
  }
  link_string<-c()
  if (!is.na(data$Social_GitHub[i])){
    if (!startsWith(data$Social_GitHub[i], "http")){
      link_string<-paste(link_string,'[{{< fa brands github >}}](https://github.com/',data$Social_GitHub[i],'){target="_blank"} ',sep="")
    }
    else {
      link_string<-paste(link_string,'[{{< fa brands github >}}](',data$Social_GitHub[i],'){target="_blank"} ',sep="")
    }
  }
  if (!is.na(data$Social_Twitter[i])){
    if (startsWith(data$Social_Twitter[i], "'@")){
      data$Social_Twitter[i] <- gsub("'@","",data$Social_Twitter[i])
    }
    if (!startsWith(data$Social_Twitter[i], "http")){
      link_string<-paste(link_string,'[{{< fa brands twitter >}}](https://twitter.com/',data$Social_Twitter[i],'){target="_blank"} ',sep="")
    }
    else {
      link_string<-paste(link_string,'[{{< fa brands twitter >}}](',data$Social_Twitter[i],'){target="_blank"} ',sep="")
    }
  }
  if (!is.na(data$Social_LinkedIn[i])){
    if (!startsWith(data$Social_LinkedIn[i], "http")){
      link_string<-paste(link_string,'[{{< fa brands linkedin >}}](https://www.linkedin.com/in/',data$Social_LinkedIn[i],'){target="_blank"} ',sep="")
    }
    else {
      link_string<-paste(link_string,'[{{< fa brands linkedin >}}](',data$Social_LinkedIn[i],'){target="_blank"} ',sep="")
    }
  }
  if (!is.na(data$Website[i])){
    if (startsWith(data$Website[i],'www')){
      data$Website[i] <- paste('http://',data$Website[i],sep="")
    }
    if (!startsWith(data$Website[i],'http')){
      data$Website[i] <- paste('http://www.',data$Website[i],sep="")
    }
    link_string<-paste(link_string,'[{{< fa globe >}}](',data$Website[i],'){target="_blank"} ',sep="")
  }
  
  methodologies<-c()
  if (length(which(method_experts[,1]==i))>0){
    for (j in 1:length(which(method_experts[,1]==i))){
      if (j>1){methodologies<-paste(methodologies,", ",sep="")}#add spacing preemptively
      if (all(colnames(data)[method_experts[which(method_experts[,1]==i),2][j]]=="Methodology-.Other") || all(colnames(data)[method_experts[which(method_experts[,1]==i),2][j]]=="Methodology")){
        methodologies<-paste(methodologies,data[i,method_experts[which(method_experts[,1]==i),2][j]+1],sep="")
      }
      else{
        methodologies<-paste(methodologies,gsub("\\."," ",gsub("Methodology-.","",
                                                               colnames(data)[method_experts[which(method_experts[,1]==i),2][j]])),sep="")
      }
      
    }
    # methodologies<-paste(methodologies,"",setp="")
  }
  else{methodologies=""}

  applications<-c()
  if (length(which(appli_res[,1]==i))){
    for (j in 1:length(which(appli_res[,1]==i))){
      if ((j>1) || (j==1 && methodologies!="")){applications<-paste(applications,", ",sep="")}
      if (all(colnames(data)[appli_res[which(appli_res[,1]==i),2][j]]=="Applications-.Other") || all(colnames(data)[appli_res[which(appli_res[,1]==i),2][j]]=="Applications")){
        applications<-paste(applications,data[i,appli_res[which(appli_res[,1]==i),2][j]+1],sep="")
      }
      else{
        applications<-paste(applications,gsub("\\."," ",gsub("Applications-.","",
                                                             colnames(data)[appli_res[which(appli_res[,1]==i),2][j]])),sep="")
      }
    }
  }
  else{applications=""}
  
  

  output_str<- paste("|",image_string,"|**",
                  name,"**<br>",data$Title[i],
                  "<br>",link_string,"|",methodologies,applications,"|",sep="")
  output_strings<-append(output_strings,gsub("; ","<br>",output_str))
}
close( file("C:/Users/Penny/Documents/DSHub/HubRoster/roster_table.txt", open="w" ) )
lapply(output_strings,write,"C:/Users/Penny/Documents/DSHub/HubRoster/roster_table.txt",append=TRUE)

# plot services results and save for website
act.data <- data.frame(
  activities = c("Networking Infrastructure","Collaboration and Match Making","Learning, Training, and Education",
                 "Disseminating Information","Advertisement of Seminars",
                 "Organization of Working Groups","Organization of Discussion Forums","External Funding Opportunities",
                 "Internal Funding Opportunities","Communication of Available Resources",
                 "Connection to Prospective Students","Industry Partnerships","Job Opportunities"),
  actcounts = c(length(grep("Networking Infrastructure",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Collaboration and Match Making",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Taking Part in Learning, Training, and Education",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Disseminating Information through Training and Educational Programs",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Advertisement of Seminars",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Organization of Working Groups",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Organization of Discussion Forums",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Communication of External Funding Opportunities",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Promotion and Offering of Internal Funding Opportunities",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Communication of Available Resources",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Connection to Prospective Students",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Industry Partnerships",data$Services))/length(which(!is.na(data$Services))),
                length(grep("Job Opportunities",data$Services))/length(which(!is.na(data$Services))))
)

p <- ggplot(data=act.data, aes(x=reorder(activities,+actcounts),
                               y=actcounts, 
                               fill=reorder(activities,+actcounts))) + 
  geom_col() + coord_flip() + ylim(0,.8) +
  geom_text(aes(label=scales::percent(actcounts)), hjust=1.5, size = 1.5) + 
  theme(legend.position = "none", axis.title = element_blank(), 
        panel.background = element_blank(), 
        axis.line = element_line(color="black"), 
        axis.text = element_text(size=6),
        axis.text.x = element_blank(),
        axis.ticks.x = element_blank())
ggsave("C:/Users/Penny/Projects/utah-data-science-hub.github.io/images/servicesPlot.png",
       p, width=1500, height=800, units = "px", dpi= 300)
  
#plot methodology summary and save for website
method.data <- data[startsWith(colnames(data),'Methodology')]
method.data <- method.data[,1:16]
colnames(method.data) <- gsub("Methodology-.","",colnames(method.data))
method.data$index <- 1:nrow(method.data)
method.data <- melt(method.data, id = "index")
method.table <- table(method.data$variable,method.data$value)
method.freq = as.data.frame(method.table)
colnames(method.freq) <- c("Methodology","Level","Freq")
method.freq = subset(method.freq,Level != "None")

p <- method.freq %>%
  mutate(Methodology = factor(Methodology, 
                              levels = arrange(method.freq[method.freq$Level=="Expert",],desc(Freq))$Methodology)) %>%
  mutate(Level = factor(Level, 
                        levels = c("Interested","User","Limited", "Intermediate", "Expert"))) %>%
  ggplot(aes(x = Methodology, y = Freq, fill = Level)) + 
  geom_bar(stat = "identity") + 
  theme(axis.title = element_blank(), 
        panel.background = element_blank(), 
        axis.line = element_line(color="black"), 
        axis.text.x = element_text(angle=90, vjust = 0.5, hjust = 1,size=8))
p
ggsave("C:/Users/Penny/Projects/utah-data-science-hub.github.io/images/methodologyPlot.png",
       p, width=1500, height=1200, units = "px", dpi= 300)


#plot applications summary and save for website
app.data <- data[startsWith(colnames(data),'Applications')]
app.data <- app.data[-(which(startsWith(colnames(app.data),"Applications-.Other")))]
colnames(app.data) <- gsub("Applications-.","",colnames(app.data))
app.data$index <- 1:nrow(app.data)
app.data <- melt(app.data, id = "index")
app.table <- table(app.data$variable,app.data$value)
app.freq = as.data.frame(app.table)
colnames(app.freq) <- c("Application","Level","Freq")
app.freq = subset(app.freq,Level != "None")

p <- app.freq %>%
  mutate(Application = factor(Application, 
                              levels = arrange(app.freq[app.freq$Level=="Developer/Researcher",],desc(Freq))$Application)) %>%
  mutate(Level = factor(Level, 
                        levels = c("Interested","User","Developer/Researcher"))) %>%
  ggplot(aes(x = Application, y = Freq, fill = Level)) + 
  geom_bar(stat = "identity") + ylim(0,50) +
  theme(axis.title = element_blank(), 
        legend.position="bottom",
        panel.background = element_blank(), 
        axis.line = element_line(color="black"), 
        axis.text.x = element_text(angle=90, vjust = 0.5, hjust = 1,size=8))
p
ggsave("C:/Users/Penny/Projects/utah-data-science-hub.github.io/images/applicationsPlot.png",
       p, width=1500, height=1200, units = "px", dpi= 300)

