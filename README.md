This is a personal project for on-the-cheap disk space monitoring.

It has two components:
- the main script "check-diskspace.ps1" which can be downloaded and run, or scheduled to run.
- the deployer script "deploy-check-diskspace.ps1" which is desinged to download the current version of the script from here on GitHub and run it daily.

*Features*
- has configurable thresholds for alerting (minimim free space and percent of space used)
- self-updates from GitHub when I release a new version
- cleans up old file versions and scheduled tasks when installed / updated.

This is a little rough, because its set up for my environment.  I am planing to add support for pulling the paramters (SMTP server, from / to address, etc) in a config file on the machine, and allow user to epcfiy their own URL to download the config file so others can use it.
