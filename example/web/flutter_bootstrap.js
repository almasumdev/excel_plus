{{flutter_js}}
{{flutter_build_config}}

const builds = Array.isArray(_flutter.buildConfig?.builds)
    ? _flutter.buildConfig.builds
    : [];

const hasSkwasmBuild = builds.some(
    (build) =>
        build.compileTarget === 'dart2wasm' && build.renderer === 'skwasm',
);

const config = hasSkwasmBuild
    ? {
        renderer: 'skwasm',
        forceSingleThreadedSkwasm: true,
      }
    : {};

_flutter.loader.load({config});