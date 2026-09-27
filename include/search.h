/*
 * Small search library used as the target of the Google Benchmark pilot.
 */

#ifndef SEARCH_H_
#define SEARCH_H_

#include <cstdint>
#include <vector>

namespace search
{
/**
 * @brief Checks if target is in values by scanning every element. O(n)
 */
bool LinearContains(const std::vector<std::int32_t>& values, std::int32_t target);

/**
 * @brief Checks if target is in values using binary search. O(log n)
 * @pre values must be sorted in ascending order
 */
bool BinaryContains(const std::vector<std::int32_t>& values, std::int32_t target);

/**
 * @brief Same as BinaryContains, but the loop body has no data-dependent branch
 *        (the comparison becomes a conditional move), avoiding branch mispredictions
 * @pre values must be sorted in ascending order
 */
bool BranchlessBinaryContains(const std::vector<std::int32_t>& values, std::int32_t target);

/**
 * @brief Generates n distinct sorted values: 0, 2, 4, ..., 2(n-1)
 */
std::vector<std::int32_t> MakeSortedEvens(std::int64_t n);

/**
 * @brief Generates count pseudo-random queries in [0, 2n) using a fixed seed
 */
std::vector<std::int32_t>
MakeQueries(std::int64_t n, std::int64_t count, std::uint32_t seed = 42);
} // namespace search

#endif // SEARCH_H_
