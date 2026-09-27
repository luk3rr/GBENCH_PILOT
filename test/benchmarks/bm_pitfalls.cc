/*
 * Classic micro-benchmarking pitfalls and the Google Benchmark tools that avoid them.
 */

#include "search.h"

#include <algorithm>
#include <benchmark/benchmark.h>
#include <list>
#include <numeric>
#include <vector>

// Pitfall 1: the result is never used, so the optimizer deletes the whole loop
static void BM_SumWithoutDoNotOptimize(benchmark::State& state)
{
    for (auto _ : state)
    {
        std::int64_t sum = 0;
        for (std::int64_t i = 0; i < state.range(0); ++i)
            sum += i * i;
    }
}
BENCHMARK(BM_SumWithoutDoNotOptimize)->Arg(1 << 16);

static void BM_SumWithDoNotOptimize(benchmark::State& state)
{
    for (auto _ : state)
    {
        std::int64_t sum = 0;
        for (std::int64_t i = 0; i < state.range(0); ++i)
        {
            sum += i * i;
            benchmark::DoNotOptimize(sum);
        }
    }
}
BENCHMARK(BM_SumWithDoNotOptimize)->Arg(1 << 16);

// Pitfall 2: same O(n) algorithm, very different constant (cache locality)
template <typename Container>
static void BM_Accumulate(benchmark::State& state)
{
    const auto n = static_cast<std::size_t>(state.range(0));
    Container  c(n);
    std::iota(c.begin(), c.end(), 0);

    for (auto _ : state)
    {
        auto sum = std::accumulate(c.begin(), c.end(), std::int64_t{0});
        benchmark::DoNotOptimize(sum);
    }

    state.SetBytesProcessed(state.iterations() * state.range(0) * sizeof(std::int32_t));
}
BENCHMARK_TEMPLATE(BM_Accumulate, std::vector<std::int32_t>)->Range(1 << 10, 1 << 20);
BENCHMARK_TEMPLATE(BM_Accumulate, std::list<std::int32_t>)->Range(1 << 10, 1 << 20);

// Pitfall 3: the input must be "fresh" at each iteration (sorting a sorted array)
static void BM_SortWithPause(benchmark::State& state)
{
    auto data = search::MakeQueries(state.range(0), state.range(0));

    for (auto _ : state)
    {
        state.PauseTiming(); // excluded from measurement (but has its own overhead!)
        auto copy = data;
        state.ResumeTiming();

        std::sort(copy.begin(), copy.end());
        benchmark::DoNotOptimize(copy.data());
        benchmark::ClobberMemory();
    }

    state.SetComplexityN(state.range(0));
}
BENCHMARK(BM_SortWithPause)
    ->RangeMultiplier(4)
    ->Range(1 << 8, 1 << 18)
    ->Unit(benchmark::kMicrosecond)
    ->Complexity(benchmark::oNLogN);
