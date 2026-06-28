# srlc_centres_data

A shared home for the team's data, scripts, and analysis. This repo keeps everything in one place so it's easy to find, reuse, and hand off — instead of files scattered across laptops, email, and chat.

If you're new to GitHub, don't worry — the [Getting started](#getting-started) section below walks you through everything you need.

## Getting started

### What is GitHub?

GitHub is an online place to store files (especially code and data) where a whole team can share the same copy, see each other's changes, and keep a full history of who changed what and when. Think of it like a shared Google Drive built specifically for working on files together — with an "undo" button that goes all the way back.

A **repo** (short for *repository*) is just one project's folder living on GitHub. This page is our repo.

### Get a copy on your computer ("cloning")

"Cloning" means downloading the repo to your own machine so you can work on it.

The easiest way if you're new is **GitHub Desktop** (a free app with buttons instead of typing commands):

1. Download and install [GitHub Desktop](https://desktop.github.com/).
2. Sign in with your GitHub account.
3. Go to **File → Clone repository**, pick `srlc_centres_data`, and choose where to save it.
4. That's it — the files are now on your computer.

If you're comfortable with the command line instead:

```bash
git clone https://github.com/<owner>/srlc_centres_data.git
cd srlc_centres_data
```

*(Replace `<owner>` with the actual account name once the repo is shared.)*

Note: You can access the command line through your preferred IDE (VS Code, etc.), through Terminal on macOS, or through Command Prompt, PowerShell, or Git Bash on Windows.

### The everyday loop

Once you have a copy, the basic rhythm is:

1. **Pull** — grab the latest changes others have made (so you're up to date).
2. **Make your changes** — edit or add files as normal.
3. **Commit** — save a snapshot of your changes with a short message describing what you did.
4. **Push** — send your changes back up to GitHub so the team can see them.

In GitHub Desktop these are all buttons; from the command line they look like:

```bash
git pull
# ...make your edits...
git add .
git commit -m "Short note about what changed"
git push
```

### A few friendly habits

- Write a short, clear commit message ("Added March centre numbers" beats "update").
- Pull before you start working so you're not building on an old copy.
- When in doubt, ask in the group — that's what it's for.

## Learning more

If you like a visual walkthrough, these are good starting points:

- GitHub's own [Quickstart guide](https://docs.github.com/en/get-started/quickstart)
- This super easy to understand [YouTube tutorial](https://www.youtube.com/watch?v=a9u2yZvsqHA)
