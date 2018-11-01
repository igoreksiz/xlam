# Creating a prerelease

```
$ git checkout master
$ npm version prerelease --preid=beta
$ npm run publish
```

# Creating a release

```
$ git checkout master
$ npm version [major|minor|patch]
$ npm run publish
