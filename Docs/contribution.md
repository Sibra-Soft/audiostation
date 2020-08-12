# Contribute fixes

This is a guide to contributing to the Audiostation open source project on Github.

### Step 1: Setup a working copy of the project

Firstly you need a local fork of the the project, so go ahead and press the “fork” button in GitHub. This will create a copy of the repository in your own GitHub account and you’ll see a note that it’s been forked underneath the project name:

![](C:\Users\Alex%20van%20den%20Berg\AppData\Roaming\marktext\images\2020-08-11-22-05-14-image.png)

Now you need a copy locally, so find the “SSH clone URL” in the right hand column and use that to clone locally using a terminal:

```batch
$ git clone git@github.com: Sibra-Soft/Audiostation.git
```

Finally, in this stage, you need to set up a new remote that points to the original project so that you can grab any changes and bring them into your local copy. Firstly clock on the link to the original repository – it’s labeled “Forked from” at the top of the GitHub page. This takes you back to the projects main GitHub page, so you can find the “SSH clone URL” and use it to create the new remote, which we’ll call *upstream*.

```batch
$ git remote add upstream git@github.com: Sibra-Soft/Audiostation.git
```

1. *origin* which points to your GitHub fork of the project. You can read and write to this remote.
2. *upstream* which points to the main project’s GitHub repository. You can only read from this remote.

### Step 2: Get it working on your machine

Now that you have the source code, get it working on your computer.

You must have a running version of Microsoft Visual Basic 6 to make changes to the project. Please use our [dependency installer](https://www.audiostation.org/downloads/dependency-installer.zip) to install all the needed dependencies

Below you will find a exmaple of a running and working Visual Basic 6 with the project

![](C:\Users\Alex%20van%20den%20Berg\AppData\Roaming\marktext\images\2020-08-11-22-16-27-image.png)

We use [CodeSmart](https://www.axtools.com/products-codesmart-vb6.php) to make programming in VB6 even more better

### Step 3: Do some work

This is the fun bit where you get to contribute to the project. It’s usually best to start by fixing a bug that is either annoying you or you’ve found on the project’s issue tracker. If you’re looking for a place to start, a lot of projects use the [“easy pick” label](http://seld.be/notes/encouraging-contributions-with-the-easy-pick-label) (or some variation) to indicate that this issue can be addressed by someone new to the project.

Now that you have picks an issue, reproduce it on your version. Once you have reproduced it, read the code to work out where the problem is. Once you’ve found the code problem, you can move on to fixing it.
