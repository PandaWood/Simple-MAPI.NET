# Releasing

Notes to self. Releasing is two steps: push a tag, then upload the package by hand.

## 1. Bump the version

`<Version>` in `src/Mapi/SimpleMapi.SDK.csproj` is the single source of truth.
Update it, and `<PackageReleaseNotes>` while you're there, then commit.

Version numbers follow [SemVer](https://semver.org): patch for a fix or a
packaging change, minor for a new API, major for anything that breaks callers.

## 2. Tag it

```
git tag 1.2.3
git push origin 1.2.3
```

Tags have no `v` prefix - just the bare number, matching the existing `1.1.0` tag.

That triggers `.github/workflows/release.yml`, which:

1. checks the tag matches `<Version>` in the csproj, and fails the build if not
2. runs `dotnet pack -c Release`
3. creates the GitHub release, with the `.nupkg` attached and release notes
   generated from the commits since the last tag

No secrets are needed - creating a release uses the automatic `GITHUB_TOKEN`.

## 3. Upload to NuGet, manually

Download the `.nupkg` from the GitHub release (or build it locally, below) and
upload it at <https://www.nuget.org/packages/manage/upload>.

This step is deliberately not automated. It would need a nuget.org API key
stored as a repo secret, which is more machinery than this project warrants.

## Building the package locally

```
dotnet pack src/Mapi/SimpleMapi.SDK.csproj -c Release
```

Lands in `artifacts/` (set by `<PackageOutputPath>`; git-ignored).

## Things that will trip you up

- **The library targets `net20;net40;netstandard2.0;net8.0-windows`.** Old
  .NET Framework consumers are served by the `net20`/`net40` assets in the
  package - that is why those targets exist, and dropping them would break
  people. It has nothing to do with what SDK you build with; the modern
  .NET SDK emits them all, on Linux too.
- **Don't set `GeneratePackageOnBuild`.** With it on, `dotnet pack` fails with
  `NU5026` on a clean tree, because packing fires per-target-framework before
  all four have built.
- **Simple MAPI is Windows-only at runtime** but nothing about building or
  packaging needs Windows, which is why CI runs on `ubuntu-latest`.
- **To actually test it** you need Windows with a default mail client
  configured. Run the console demo in `src/Demos/Console-Demo`. There are no
  automated tests - the whole library is P/Invoke into `MAPI32.DLL`.
