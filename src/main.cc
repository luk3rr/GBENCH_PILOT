#include "search.h"

#include <cstdlib>
#include <iostream>

int main()
{
    const auto values = search::MakeSortedEvens(10);

    std::cout << "LinearContains(8) = " << search::LinearContains(values, 8) << '\n'
              << "BinaryContains(7) = " << search::BinaryContains(values, 7) << std::endl;

    return EXIT_SUCCESS;
}
