# Zaitounish

Mohamed Zaitoun's portfolio and resume live at [www.mohamedzaitoun.com](https://www.mohamedzaitoun.com/).

## Update the portfolio and resume

Use GitHub’s editor to update contact details, experience, education, skills and projects in one file. The build uses those records for the website, PDF and editable Word resume. [Edit the profile](https://github.com/zaitounish/MohamedZaitoun/edit/main/content/profile.json)

The published PDF always comes from the latest successful build. [Download the current resume](https://www.mohamedzaitoun.com/Mohamed%20Zaitoun%20Resume.pdf)

From the repository root, use Python 3.12 to build, check and serve a local preview.

```sh
python -m pip install -r requirements.txt
python scripts/build.py --output dist
python scripts/check.py --output dist
python -m http.server 8000 --bind 127.0.0.1 --directory dist/site
```

Open [the local preview](http://127.0.0.1:8000/). The public download is `dist/site/Mohamed Zaitoun Resume.pdf`. The editable version is `dist/resume/Mohamed Zaitoun Resume.docx`. The optional `--word-copy` flag saves another copy to the filename you specify.

```sh
python scripts/build.py --output dist --word-copy "../Mohamed Zaitoun Resume.docx"
```

## Publish

Open a pull request. In **Actions > Portfolio**, the build checks the generated files and attaches the editable Word resume and the packaged site. A successful push to `main`, or a manual run on `main`, publishes `dist/site`.

The Pages publishing source is **GitHub Actions**. Keep the custom domain `www.mohamedzaitoun.com` in the repository’s Pages settings and its `CNAME` file. The publishable homepage is the generated `dist/site/index.html`. [GitHub Pages publishing settings](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site)
