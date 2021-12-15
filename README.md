# Coloma
Event Log Grabber

Coloma grabs the local event log and saves it to a TSV. This was for an internal project where I collected event logs from multiple devices to find ways to improve Windows, add additional telemetry, etc. @stanleyhon worked on this with me.

The interesting piece we discovered is that when a PC has hardware that is failing, it can detect certain failures.

There are definitely quirks, for example the TSV file saved by Coloma is saved in poorly chosen folder, and although the file is readable in Excel it's not ideal. A custom viewer would be better. I don't know what ColomaAnalysis does, but @stanleyhon might.
