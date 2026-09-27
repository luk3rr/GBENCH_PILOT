#define DOCTEST_CONFIG_IMPLEMENT_WITH_MAIN
#include "doctest.h"
#include "search.h"

TEST_CASE("Linear and binary search agree")
{
    const auto values  = search::MakeSortedEvens(1000);
    const auto queries = search::MakeQueries(1000, 1000);

    for (const auto q : queries)
    {
        CHECK(search::LinearContains(values, q) == search::BinaryContains(values, q));
        CHECK(search::LinearContains(values, q) == search::BranchlessBinaryContains(values, q));
    }
}

TEST_CASE("Edge cases")
{
    const std::vector<std::int32_t> empty;
    CHECK_FALSE(search::BinaryContains(empty, 0));
    CHECK_FALSE(search::LinearContains(empty, 0));
    CHECK_FALSE(search::BranchlessBinaryContains(empty, 0));

    const auto values = search::MakeSortedEvens(5); // 0 2 4 6 8
    CHECK(search::BinaryContains(values, 0));
    CHECK(search::BinaryContains(values, 8));
    CHECK_FALSE(search::BinaryContains(values, 9));
    CHECK_FALSE(search::BinaryContains(values, -1));

    for (std::int32_t q = -1; q <= 9; ++q)
        CHECK(search::BranchlessBinaryContains(values, q) == search::LinearContains(values, q));
}
