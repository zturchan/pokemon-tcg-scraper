#-------------------------------------------------------------------------------
# Name:        Pokemon TCG Scraper
# Purpose:     Scrapes the POKEMON TCG API for all cards and renders them in
#              a desired Excel format.
#
# Author:      Zak Turchansky
#
# Created:     15-06-2019
# Copyright:   (c) Zak Turchansky 2019
# Licence:     LICENSE.txt
#-------------------------------------------------------------------------------

from pokemontcgsdk import Card

def main():
    fields_remain_same = ['id',
                          'name',
                          'national_pokedex_number',
                          'image_url',
                          'image_url_hi_res',
                          'subtype',
                          'supertype',
                          'ancient_trait',
                          'hp',
                          'number',
                          'artist',
                          'rarity',
                          'series',
                          'set',
                          'set_code',
                          'converted_retreat_cost',
                          'types',
                          'evolves_from']

if __name__ == '__main__':
    main()
