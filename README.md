# excel-azure-ai-translator
 Pour traduire des colonnes spécifiques d'une feuille de travail Excel dans plusieurs langues à l'aide du module de traduction basé sur l'intelligence artificielle Azure AI Translator 

| Langue                       | Code langue | Langue                         | Code langue | Langue                         | Code langue | Langue                         | Code langue |
|------------------------------|:-----------:|--------------------------------|:-----------:|--------------------------------|:-----------:|--------------------------------|:-----------:|
| Afrikaans                    | `af`        | Italien                        | `it`        | Albanais                       | `sq`        | Japonais                       | `ja`        |
| Amharique                    | `am`        | Kannada                        | `kn`        | Arabe                          | `ar`        | Kashmiri                       | `ks`        |
| Arménien                     | `hy`        | Kazakh                         | `kk`        | Assamais                       | `as`        | Khmer                          | `km`        |
| Azerbaïdjanais (Latin)       | `az`        | Kinyarwanda                    | `rw`        | Bangla                         | `bn`        | Klingon                        | `tlh-Latn`  |
| Bashkir                      | `ba`        | Klingon (plqaD)                | `tlh-Piqd`  | Basque                         | `eu`        | Konkani                        | `gom`       |
| Bhojpouri                    | `bho`       | Coréen                         | `ko`        | Bodo                           | `brx`       | Kurde (central)                | `ku`        |
| Bosniaque (latin)            | `bs`        | Kurde (Nord)                   | `kmr`       | Bulgare                        | `bg`        | Kirghiz (cyrillique)           | `ky`        |
| Cantonais (traditionnel)     | `yue`       | Lao                            | `lo`        | Catalan                        | `ca`        | Letton                         | `lv`        |
| Chinois (littéraire)         | `lzh`       | Lituanien                      | `lt`        | Chinois (simplifié)            | `zh-Hans`   | Lingala                        | `ln`        |
| Chinois traditionnel         | `zh-Hant`   | Bas sorabe                     | `dsb`       | chiShona                       | `sn`        | Luganda                        | `lug`       |
| Croate                       | `hr`        | Macédonien                     | `mk`        | Tchèque                        | `cs`        | Maithili                       | `mai`       |
| Danois                       | `da`        | Malgache                       | `mg`        | Dari                           | `prs`       | Malais (latin)                 | `ms`        |
| Maldivien                    | `dv`        | Malayalam                      | `ml`        | Dogri                          | `doi`       | Maltais                        | `mt`        |
| Néerlandais                  | `nl`        | Maori                          | `mi`        | Anglais                        | `en`        | Marathi                        | `mr`        |
| Estonien                     | `et`        | Mongole (cyrillique)           | `mn-Cyrl`   | Féroïen                        | `fo`        | Mongol (traditionnel)          | `mn-Mong`   |
| Fidjien                      | `fj`        | Myanmar                        | `my`        | Filipino                       | `fil`       | Népalais                       | `ne`        |
| Finnois                      | `fi`        | Norvégien                      | `nb`        | Français                       | `fr`        | Nyanja                         | `nya`       |
| Français (Canada)            | `fr-ca`     | Odia                           | `or`        | Galicien                       | `gl`        | Pachto                         | `ps`        |
| Géorgien                     | `ka`        | Persan                         | `fa`        | Allemand                       | `de`        | Polonais                       | `pl`        |
| Grec                         | `el`        | Portugais (Brésil)             | `pt`        | Goudjrati                      | `gu`        | Portugais (Portugal)           | `pt-pt`     |
| Créole haïtien               | `ht`        | Pendjabi                       | `pa`        | Hausa                          | `ha`        | Queretaro Otomi                | `otq`       |
| Hébreu                       | `he`        | Roumain                        | `ro`        | Hindi                          | `hi`        | Rundi                          | `run`       |
| Hmong daw (latin)            | `mww`       | Russe                          | `ru`        | Hongrois                       | `hu`        | Samoan (latin)                 | `sm`        |
| Islandais                    | `is`        | Serbe (cyrillique)             | `sr-Cyrl`   | Igbo                           | `ig`        | Serbe (latin)                  | `sr-Latn`   |
| Indonésien                   | `id`        | Sesotho                        | `st`        | Inuinnaqtun                    | `ikt`       | Sotho du Nord                  | `nso`       |
| Inuktitut                    | `iu`        | Setswana                       | `tn`        | Inuktitut (Latin)              | `iu-Latn`   | Sindhi                         | `sd`        |
| Irlandais                    | `ga`        | Cingalais                      | `si`        | Italien                        | `it`        | Slovaque                       | `sk`        |
| Japonais                     | `ja`        | Slovène                        | `sl`        | Kannada                        | `kn`        | Somali (arabe)                 | `so`        |
| Kashmiri                     | `ks`        | Espagnol                       | `es`        | Kazakh                         | `kk`        | Swahili (latin)                | `sw`        |
| Khmer                        | `km`        | Suédois                        | `sv`        | Kinyarwanda                    | `rw`        | Tahitien                       | `ty`        |
| Klingon                      | `tlh-Latn`  | Tamoul                         | `ta`        | Klingon (plqaD)                | `tlh-Piqd`  | Tatar (latin)                  | `tt`        |
| Konkani                      | `gom`       | Télougou                       | `te`        | Coréen                         | `ko`        | Thaï                           | `th`        |
| Kurde (central)              | `ku`        | Tibétain                       | `bo`        | Kurde (Nord)                   | `kmr`       | Tigrigna                       | `ti`        |
| Kirghiz (cyrillique)         | `ky`        | Tonga                          | `to`        | Lao                            | `lo`        | Turc                           | `tr`        |
| Letton                       | `lv`        | Turkmène (latin)               | `tk`        | Lituanien                      | `lt`        | Ukrainien                      | `uk`        |
| Lingala                      | `ln`        | Haut sorabe                    | `hsb`       | Bas sorabe                     | `dsb`       | Ourdou                         | `ur`        |
| Luganda                      | `lug`       | Ouïgour (arabe)                | `ug`        | Macédonien                     | `mk`        | Ouzbek (latin)                 | `uz`        |
| Maithili                     | `mai`       | Vietnamien                     | `vi`        | Malgache                       | `mg`        | Gallois                        | `cy`        |
| Malais (latin)               | `ms`        | Xhosa                          | `xh`        | Malayalam                      | `ml`        | Yoruba                         | `yo`        |
| Maltais                      | `mt`        | Maori                          | `mi`        | Malayalam                      | `ml`        | Zoulou                         | `zu`        |

