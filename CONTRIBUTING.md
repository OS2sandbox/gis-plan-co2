# Make contributing easy

It's important that it is very clear how you can become involved and contribute to a project.

Therefor:

* The project MUST have a public issue tracker that accepts suggestions from anyone. In OS2 we use [Jira](https://os2web.atlassian.net/) but a combination with for example Github issues are also okay.
* The project MUST include instructions for how to privately report security issues for responsible disclosure. This can be covered by the [SECURITY](SECURITY.md) file.
* The documentation MUST link to both the public issue tracker and submitted codebase changes, for example in a README file.
* The project MUST have communication channels for users and developers. In OS2 we provide the tools to communicate through email lists and our main website.
* The documentation SHOULD include instructions for how to report potentially security sensitive issues on a closed channel.

# Example

# Contribution guidelines for *OS2produkt*

If you want to contribute to *OS2produkt*, we ask you to follow these guidelines.

## Reporting bugs
If you have encountered a bug in *OS2produkt*, please check if an issue already exists in the list of existing [issues](https://os2web.atlassian.net/), if such an issue does not exist, you can create one [here](https://os2web.atlassian.net/). When writing the bug report, try to add a clear example that shows how to reproduce said bug.

## Adding new features
Before making making changes to the code, we advise you to first check the list of existing [issues](https://os2web.atlassian.net/) for OS2produkt to see if an issue for the suggested changes already exists. If such an issue does not exist, you can create one [here](https://os2web.atlassian.net/). Creating an issue gives an opportunity for other developers to give tips even before you start coding. If you are in the early idea phase, or if your feature requires larger changes, you can discuss it with the [projects coordination group](https://os2.eu) or by [contacting OS2 directly](https://os2.eu/kontakt) to make sure you are heading in the right direction.

### Code style
To keep the code clean and readable, *Please include links to coding standards and tools used:*

Whenever a branch is pushed or a pull request is made, the code will be checked in CI by the tools mentioned above, so make sure to install these tools and run them locally before pushing branches/making PRs.

### Forking the repository
In order to implement changes to *OS2produkt* when you do not have rights for the required repository, you must first fork the repository. Once the repository is forked, you can clone it to your local machine.

### Making the changes
On your local machine, create a new branch, and name it like:
- `feature/some-new-feature`, if the changes implement a new feature
- `issue/some-issue`, if the changes fix an issue

Once you have made changes or additions to the code, you can commit them (try to keep the commit message descriptive but short). If an issue exists in the issue list for the changes you made, be sure to format your commit message like `"Fixes #<issue_id> -- description of changes made`, where `<issue_id>"` corresponds to the number of the issue on Jira or GitHub. To demonstrate that the changes implement the new feature/fix the issue, make sure to also add tests.

### Making a pull request
If all changes have been committed, you can push the branch to your fork of the repository and create a pull request to the `master` branch of the *OS2produkt* repository. Your pull request will be reviewed, if applicable feedback will be given and if everything is approved, it will be merged.

### Reviews on releases

All pull requests will be reviewed before they are merged to a release branch. As well as being reviewed for functionality and following the code style they will be checked with the steering committee and/or the coordination group for the project.
