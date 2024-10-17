#!/usr/bin/env python3


from src.equity_equation.equity_equation_investors import EquityEquation

if __name__ == "__main__":
    eeq = EquityEquation(
        total_equity_pool=float(
            input("""Enter the total equity pool as a percentage (total_equity_pool, between 0 and 100%): """)
        ),
        proposed_equity_alloc=float(
            input(
                """Enter the proposed equity percentage as a
                   percentage to be given to the investor as a result
                   of their investment in the company
                   (proposed_equity_alloc, between 0 and 100%): """
            )
        ),
        estimated_valuation_increase=float(
            input(
                """Enter the estimated increase in valuation of the company
                   as a result of the investor investing,
                   as a percentage. Said differently,
                   how much you expect the valuation of the
                   company to inrease after their investment
                   (estimated_valuation_increase): """
            )
        ),
        pre_money_valuation=float(
            input(
                """Enter the valuation of the company before the investment (the pre-money valuation)
                   (pre_money_valuation): """
            )
        ),
    )
    eeq()
