using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Toolchains.InProcess.NoEmit;

var config = DefaultConfig.Instance
    .AddJob(Job.Default
        .WithToolchain(InProcessNoEmitToolchain.Instance));

BenchmarkSwitcher.FromTypes([
    typeof(StreamScheme.Benchmark.WriteUniqueStrings),
    typeof(StreamScheme.Benchmark.WriteSparseCategories),
    typeof(StreamScheme.Benchmark.WriteMixedData),
    typeof(StreamScheme.Benchmark.StreamingRead),
]).Run(args, config);
