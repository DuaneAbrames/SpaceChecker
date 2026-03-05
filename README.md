This is a personal project for on-the-cheap disk space monitoring.

It has two components:
- the main script "check-diskspace.ps1" which can be downloaded and run, or scheduled to run.
- the deployer script "deploy-check-diskspace.ps1" which is desinged to download the current version of the script from here on GitHub and schedule it to run daily.

*Features*
- has configurable thresholds for alerting (minimim free space and percent of space used)
- self-updates from GitHub when I release a new version
- cleans up old file versions and scheduled tasks when installed / updated.
- I use ScreenConnect, so I have added the #!ps and #timeout directly into the deployment script for easy copy-pasting. (IYKYK)

Configuration Files:
You can do the following:
1. distribute the config file manually via whatever method you choose
2. hard-code the URL during the deploy script run either from command line, or editing the script before runtim
3. the script will fallback to a url constructed from "space" and the primary DNS domain of the machine (eg https://space.example.org)

I have provided an example config file, you will need to edit it for your site-specific configuration (SMTP)