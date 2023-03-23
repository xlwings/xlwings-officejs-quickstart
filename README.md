Elastic Beanstalk deployment
----------------------------

1. Install EB CLI: https://github.com/aws/aws-elastic-beanstalk-cli-setup
3. `eb init -p docker <app name>` (use any name for `<app name>`)
4. `eb create <env name>` (use any name for `<env name>`)
3. `eb setenv XLWINGS_LICENSE_KEY=<your-license-key>`
5. Open the URL: `eb open`, you should see `{"status":"ok"}`
6. Configure SSL: in Elastic Beanstalk, go to you Environment, then click on Configuration > Load Balancer > Edit. Add a Listener at the top, as explained in this video:
https://www.youtube.com/watch?v=pBOcPxho_wg (you'll require a custom domain and an SSL certificate)

Deploy updates:

* Commit the changes to Git then run: `eb deploy`
* See also https://docs.aws.amazon.com/elasticbeanstalk/latest/dg/eb3-cli-git.html

Logs:

* Get logs via `eb logs`

Remove everything:

* `eb terminate <env name>`
* Note that the corresponding S3 bucket is not deleted automatically.
