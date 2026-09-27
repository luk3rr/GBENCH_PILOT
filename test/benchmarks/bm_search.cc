/*
 * Linear vs binary search vs hash lookup, with asymptotic complexity fitting.
 */

#include "search.h"

#include <benchmark/benchmark.h>
#include <unordered_set>

// Always use many queries: with only N queries cycling, the branch predictor
// memorizes the pattern for small N and the numbers become too optimistic
constexpr std::int64_t kNumQueries = 1 << 16;

// 1) Simplest possible benchmark: setup outside the loop, only the body is timed
static void BM_LinearSearch(benchmark::State& state)
{
    const auto values  = search::MakeSortedEvens(state.range(0));
    const auto queries = search::MakeQueries(state.range(0), kNumQueries);

    std::size_t i = 0;
    for (auto _ : state)
    {
        bool found = search::LinearContains(values, queries[i]);
        benchmark::DoNotOptimize(found);
        i = (i + 1) % queries.size();
    }

    state.SetItemsProcessed(state.iterations());
    state.SetComplexityN(state.range(0));
}
BENCHMARK(BM_LinearSearch)->RangeMultiplier(4)->Range(16, 1 << 16)->Complexity();

static void BM_BinarySearch(benchmark::State& state)
{
    const auto values  = search::MakeSortedEvens(state.range(0));
    const auto queries = search::MakeQueries(state.range(0), kNumQueries);

    std::size_t i = 0;
    for (auto _ : state)
    {
        bool found = search::BinaryContains(values, queries[i]);
        benchmark::DoNotOptimize(found);
        i = (i + 1) % queries.size();
    }

    state.SetItemsProcessed(state.iterations());
    state.SetComplexityN(state.range(0));
}
BENCHMARK(BM_BinarySearch)->RangeMultiplier(4)->Range(16, 1 << 16)->Complexity();

static void BM_BranchlessBinarySearch(benchmark::State& state)
{
    const auto values  = search::MakeSortedEvens(state.range(0));
    const auto queries = search::MakeQueries(state.range(0), kNumQueries);

    std::size_t i = 0;
    for (auto _ : state)
    {
        bool found = search::BranchlessBinaryContains(values, queries[i]);
        benchmark::DoNotOptimize(found);
        i = (i + 1) % queries.size();
    }

    state.SetItemsProcessed(state.iterations());
    state.SetComplexityN(state.range(0));
}
BENCHMARK(BM_BranchlessBinarySearch)->RangeMultiplier(4)->Range(16, 1 << 16)->Complexity();

// 2) Fixture: shared setup/teardown for a family of benchmarks
class HashFixture : public benchmark::Fixture
{
  public:
    void SetUp(const benchmark::State& state) override
    {
        const auto values = search::MakeSortedEvens(state.range(0));
        set               = std::unordered_set<std::int32_t>(values.begin(), values.end());
        queries           = search::MakeQueries(state.range(0), kNumQueries);
    }

    void TearDown(const benchmark::State&) override
    {
        set.clear();
        queries.clear();
    }

  protected:
    std::unordered_set<std::int32_t> set;
    std::vector<std::int32_t>        queries;
};

BENCHMARK_DEFINE_F(HashFixture, BM_HashLookup)(benchmark::State& state)
{
    std::size_t  i    = 0;
    std::int64_t hits = 0;

    for (auto _ : state)
    {
        bool found = set.contains(queries[i]);
        benchmark::DoNotOptimize(found);
        hits += found;
        i = (i + 1) % queries.size();
    }

    state.SetItemsProcessed(state.iterations());
    state.SetComplexityN(state.range(0));

    // 3) Custom counter: shown as an extra column in the output
    state.counters["hit_rate"] =
        benchmark::Counter(static_cast<double>(hits) / static_cast<double>(state.iterations()));
}
BENCHMARK_REGISTER_F(HashFixture, BM_HashLookup)
    ->RangeMultiplier(4)
    ->Range(16, 1 << 16)
    ->Complexity(benchmark::o1);
