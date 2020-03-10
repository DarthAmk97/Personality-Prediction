ny <- read.csv("D:/AI/nytimes_facebook_statuses.csv")
testing=boxplot(ny$num_reactions)
outliers_reactions = boxplot(ny$num_reactions)$out #getting outliers for reactions
outliers_likes = boxplot(ny$num_likes)$out #getting outliers for likes.

copy = ny # making copy of originar data for manipulation

for(i in outliers_reactions){ #setting NA to outliers in reactions
  copy$num_reactions[copy$num_reactions==i]<-NA
  copy$num_reactions[copy$num_reactions==i]<-NA
  
}
for(j in outliers_likes){ #setting NA to outliers in likes
  copy$num_likes[copy$num_likes==j]<-NA
  copy$num_likes[copy$num_likes==j]<-NA
}

copy <- na.omit(copy) #omitting all rows having NA from data

set.seed(50)  # setting seed to reproduce results of random sampling
trainingRowIndex <- sample(1:nrow(copy), 0.7*nrow(copy))  # row indices for training data
trainingData <- copy[trainingRowIndex, ]  # model training data
testData  <- copy[-trainingRowIndex, ]   # test data

linearModel <- lm( num_reactions ~ num_likes, data=trainingData)  # build the model
reactionsPredict <- predict(linearModel, testData)  # pre
