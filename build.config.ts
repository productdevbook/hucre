import { defineBuildConfig } from "obuild/config"

export default defineBuildConfig({
  entries: [
    {
      type: "transform",
      input: "./src",
      outDir: "./dist",
      dts: true,
      // The CLI is bundled separately (below) so its dependencies land in
      // the shipped file. Transforming it here as well would emit a second
      // dist/cli.mjs carrying bare `citty` / `consola` imports — which is
      // exactly how the published binary ended up broken. See #357.
      //
      // `src/cli/` is excluded for the same reason: the bin is a two-line
      // wrapper around ./cli/commands (#399), and transforming that module
      // would put the unresolvable bare imports back in dist under a new
      // name. The bundle entry inlines it into dist/cli.mjs instead.
      filter: (filePath) => !/(?:^|[/\\])cli(?:\.ts$|[/\\])/.test(filePath),
    },
    {
      // The library keeps zero runtime dependencies, so the CLI's own deps
      // have to be baked into the binary rather than resolved from
      // node_modules at run time.
      type: "bundle",
      input: "./src/cli.ts",
      outDir: "./dist",
      dts: false,
      minify: true,
    },
  ],
})
