#include "search.h"

#include <random>

namespace search
{
bool LinearContains(const std::vector<std::int32_t>& values, std::int32_t target)
{
    for (const auto v : values)
    {
        if (v == target)
            return true;
    }

    return false;
}

bool BinaryContains(const std::vector<std::int32_t>& values, std::int32_t target)
{
    std::size_t lo = 0;
    std::size_t hi = values.size();

    while (lo < hi)
    {
        const std::size_t mid = lo + (hi - lo) / 2;

        if (values[mid] < target)
            lo = mid + 1;
        else
            hi = mid;
    }

    return lo < values.size() && values[lo] == target;
}

bool BranchlessBinaryContains(const std::vector<std::int32_t>& values, std::int32_t target)
{
    if (values.empty())
        return false;

    const std::int32_t* base = values.data();
    std::size_t         len  = values.size();

    while (len > 1)
    {
        const std::size_t half = len / 2;
        base                   = (base[half] < target) ? base + half : base;
        len -= half;
    }

    base += (*base < target);

    return base < values.data() + values.size() && *base == target;
}

std::vector<std::int32_t> MakeSortedEvens(std::int64_t n)
{
    std::vector<std::int32_t> values(static_cast<std::size_t>(n));

    for (std::int64_t i = 0; i < n; ++i)
        values[static_cast<std::size_t>(i)] = static_cast<std::int32_t>(2 * i);

    return values;
}

std::vector<std::int32_t>
MakeQueries(std::int64_t n, std::int64_t count, std::uint32_t seed)
{
    std::mt19937                                gen(seed);
    std::uniform_int_distribution<std::int32_t> dist(0, static_cast<std::int32_t>(2 * n - 1));

    std::vector<std::int32_t> queries(static_cast<std::size_t>(count));

    for (auto& q : queries)
        q = dist(gen);

    return queries;
}
} // namespace search
