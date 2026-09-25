import sys
from pathlib import Path
import unittest

ROOT=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(ROOT/'src'))
from allocations import ai_answer, calculated_values, clean_record, parse_ai_json, row_values


class AllocationTests(unittest.TestCase):
    def test_note_card_fields_become_simulation_values(self):
        record=clean_record({
            'reference_number':'IIF2113-17092026','company_id':6001,
            'issuer_name':'Precision Tools Sdn Bhd','note_name':'Precision Tools Manufacturer 11',
            'product_type':'Islamic Invoice Financing (IIF) - Receivables','rating':'CR6',
            'business_description':'Precision tool manufacturing','investment_amount':'63729.67',
            'loan_note_size':'386000','allocation_date':'2026-09-17','disbursal_date':'2026-09-25',
            'payment_type':'Profit Only','term':'90','tenor_type':'Days','gross_pa':'15.60',
            'campaign_start':'2026-09-17','campaign_end':'2026-09-24','status':'Active'
        })
        self.assertEqual(record['loan_code'],'IIF-2113')
        self.assertEqual(record['gross_pa'],.156)
        values=calculated_values(record)
        self.assertEqual(values['Final Repayment Date '],'2026-12-24')
        self.assertAlmostEqual(values['Investment Exposure'],63729.67/386000)
        self.assertAlmostEqual(values['Service Fee '],values['Gross Profit']*.2)
        headers=['Gross Profit Earned','Total Gross Profit','Late Payment Charges','Service Fee ','SST','Net Profit','Installment']
        expanded=dict(zip(headers,row_values(headers,record)))
        self.assertEqual(expanded['Gross Profit Earned'],expanded['Total Gross Profit'])
        self.assertEqual(expanded['Late Payment Charges'],0)
        self.assertEqual(expanded['SST'],0)
        self.assertAlmostEqual(expanded['Net Profit'],expanded['Total Gross Profit']-expanded['Service Fee '])
        self.assertAlmostEqual(expanded['Installment'],expanded['Gross Profit Earned']/90)

    def test_fenced_ai_response_is_parsed_without_extra_fields(self):
        result=parse_ai_json('''```json
        {"note_name":"Precision Tools Manufacturer 11","reference_number":"IIF2113-17092026",
        "financing_amount":386000,"unexpected":"discard"}
        ```''')
        self.assertEqual(result['note_name'],'Precision Tools Manufacturer 11')
        self.assertEqual(result['financing_amount'],386000)
        self.assertNotIn('unexpected',result)

    def test_workers_ai_result_envelope_and_percentage_are_normalized(self):
        response={'result':{'answer':'{"profit_rate_pa":0.156,"note_tenure":90,"tenor_type":null}'}}
        result=parse_ai_json(ai_answer(response))
        self.assertEqual(result['profit_rate_pa'],15.6)
        self.assertEqual(result['tenor_type'],'Days')

    def test_allocation_cannot_exceed_financing_amount(self):
        with self.assertRaisesRegex(ValueError,'cannot exceed'):
            clean_record({
                'loan_code':'IIF-9999','company_id':1,'issuer_name':'Issuer','note_name':'Note',
                'product_type':'Islamic Invoice Financing (IIF) - Receivables','business_description':'Business',
                'investment_amount':101,'loan_note_size':100,'allocation_date':'2026-09-25',
                'disbursal_date':'2026-09-26','payment_type':'Bullet','term':90,
                'tenor_type':'Days','gross_pa':10
            })


if __name__=='__main__': unittest.main()
