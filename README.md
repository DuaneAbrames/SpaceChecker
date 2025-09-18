This is a personal project for on-the-cheap disk space monitoring.

It has two components:
- the main script "check-diskspace.ps1" which can be downloaded and run, or scheduled to run.
- the deployer script "deploy-check-diskspace.ps1" which is desinged to download the current version of the script from here on GitHub and run it daily.

*Features*
- has configurable thresholds for alerting (minimim free space and percent of space used)
- self-updates from GitHub when I release a new version
- cleans up old file versions and scheduled tasks when installed / updated.
- I use ScreenConnect, so I have added the #!ps and #timeout directly into the deployment script for easy copy-pasting.

(WIP): The config file can be overridden for your environment by spefifying a URL in the CheckDiskSpace environment variable.  If that URL is set, then every time I update the script, it will pull down a new copy of the config.  If you do not set the environment variable, then you can distribute your config files in any way you choose.
