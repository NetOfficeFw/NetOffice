# NetOffice repository history

Original NetOffice project is located at [http://netoffice.codeplex.com/](http://netoffice.codeplex.com/).
Over the years, NetOffice became a very big repository and Codeplex became very slow to work with it.
It is problematic to checkout the source code, view messages log or even do simple updates, as SvnBridge
is not working properly with NetOffice repository at Codeplex (problems included timeouts with Subversion
operations, Subversion is not able to export all files from a revision from time to time and viewing logs
or diffs with previous revisions is too long).

This Git repository was created by importing each individual revision from Subversion repository into Git.
Some revisions were omitted from import and some were were modified to exclude binary files (eg. build
assemblies, some RAR files) which were unnecessary and just consumed a lot of space.


## Excluded revisions

These revisions from original Subversion repository were not imported, as they changeset contained only
deletion of all source code files.

* [87527](https://netoffice.codeplex.com/SourceControl/changeset/87527)
* [87535](https://netoffice.codeplex.com/SourceControl/changeset/87535)
* [89344](https://netoffice.codeplex.com/SourceControl/changeset/89344)
* [90048](https://netoffice.codeplex.com/SourceControl/changeset/90048)
* [90432](https://netoffice.codeplex.com/SourceControl/changeset/90432)
* [93187](https://netoffice.codeplex.com/SourceControl/changeset/93187)
* [94556](https://netoffice.codeplex.com/SourceControl/changeset/94556)
* [103898](https://netoffice.codeplex.com/SourceControl/changeset/103898)
* [104640](https://netoffice.codeplex.com/SourceControl/changeset/104640)
