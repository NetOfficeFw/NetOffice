# NetOfficeFw/NetOffice Tepository History

Original NetOffice project was started at [http://netoffice.codeplex.com/](http://netoffice.codeplex.com/).
Over the years, NetOffice became a very big repository and Codeplex became very slow to work with it.
In 2016 the SvnBridge used by Codeplex stopped working so the original repository is not accessible anymore.

This Git repository was created by importing each individual revision from Subversion repository into Git.
Some revisions were omitted from import and some were were modified to exclude binary files (eg. built
assemblies, some RAR files) which were unnecessary and just consumed a lot of space.

The imported history can be found in the `import/legacy_repository` branch.

## Excluded revisions

These revisions were not imported from original Subversion repository, as those changesets contained only
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
