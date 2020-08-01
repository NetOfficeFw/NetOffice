# NetOfficeFw/NetOffice Repository History

Original NetOffice project was started at [http://netoffice.codeplex.com/](http://netoffice.codeplex.com/).
Over the years, NetOffice became a very big repository and Codeplex was having hard time serving
source code from this large repository.
In 2016 the SvnBridge used by Codeplex stopped working so the original repository is not accessible anymore.


## Imported commits from Subversion

This Git repository was created by importing each individual revision from Subversion repository into Git.
Some revisions were omitted from import and some were modified to exclude binary files (eg. built
assemblies, some RAR files) which were unnecessary and consumed a lot of space.

> I was lucky and I downloaded all Subversion commits before the SvnBridge stopped working.
> These individually checked out revisions were imported to Git one-by-one to create
> new Git repository.

The original imported history can be found in the `import/legacy_repository` branch.

This imported history is different from the legacy [NetOffice](https://github.com/netoffice/NetOffice) repository.
The legacy repository includes all the binary files and also commits which deleted all the files in the project.


## NetOffice 1.7.4

Sebastian started to work on 1.7.4 in summer 2017. He made the source code available to me
by sending me ZIP files which I was adding to the `import/netoffice_1.7.4-alpha` branch.


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
