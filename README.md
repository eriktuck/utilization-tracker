An app to help EI staff manage towards their utilization target.

Find the app at https://ei-utilization.herokuapp.com.

Data is updated daily around 9a ET/7 MT/6 PT.



**To maintain**:

Data from the Scoreboard should be pasted into the Utilization Input datasheet, 'TARGETS' worksheet, each month.

**To deploy changes:**

1. Commit changes to github

   `git commit -am "<message>"`

   `git pull`

   `git push origin master`

2. Push changes heroku:

   `git push heroku master`

