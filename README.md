# Excel Azure AI Translator

[![Language](https://img.shields.io/badge/.NET-8.0.101-Language?color=blue)](https://dotnet.microsoft.com)
[![Nug](https://img.shields.io/badge/EPPlus-7.0.9-Module)](https://www.nuget.org/packages/EPPlus/7.0.9)
[![Nug](https://img.shields.io/badge/Newtonsoft.Json-13.0.3-Module)](https://www.nuget.org/packages/Newtonsoft.Json/13.0.3)
[![Nug](https://img.shields.io/badge/DotNetEnv-3.0.0-Module)](https://www.nuget.org/packages/DotNetEnv/3.0.0)
[![Api](https://img.shields.io/badge/Azure%20AI%20Translator-3.0-Api?color=yellow)](https://learn.microsoft.com/fr-fr/azure/ai-services/translator/)
[![Copyright](https://img.shields.io/badge/Cr%C3%A9ateur-Kiplay03-Copyright?color=red)](https://github.com/Kiplay03)

The Excel Azure AI Translator project has been designed to massively translate the cells of an Excel column from one language to another in a fully automated way. I opted to use [Azure's AI-based translation service](https://azure.microsoft.com/en-us/products/ai-services/ai-translator) because of its free plan offering a monthly translation allowance of 2 million characters.

## Summary

- [What languages ​​does the program support?](#supportedLanguages)
- [How to install the program?](#install)
- [How to format the Excel file to provide to the program?](#formatExcelFile)
- [How to start the program?](#start)
- [How to improve translation accuracy with Custom Translator?](#improveAccuracyWithCustomTranslator)

## Supported languages <a name="supportedLanguages"></a>

Below you can find the languages ​​supported by the program and their associated language codes to properly format your Excel file and provide the program with a valid target language or language code for translation.

| Language                  | Language code | Language                  | Language code | Language                  | Language code |
|---------------------------|:-------------:|---------------------------|:-------------:|---------------------------|:-------------:|
| Afrikaans                 | `af`          | Hungarian                 | `hu`          | Polish                    | `pl`          |
| Albanian                  | `sq`          | Icelandic                 | `is`          | Portuguese (Brazil)		| `pt`          |
| Amharic                   | `am`          | Igbo                      | `ig`          | Portuguese (Portugal)		| `pt-pt`       |
| Arabic                    | `ar`          | Indonesian                | `id`          | Punjabi                   | `pa`          |
| Armenian                  | `hy`          | Inuinnaqtun               | `ikt`         | Queretaro Otomi			| `otq`         |
| Assamese                  | `as`          | Inuktitut                 | `iu`          | Romanian                  | `ro`          |
| Azerbaijani (Latin)       | `az`          | Inuktitut (Latin)         | `iu-Latn`     | Rundi                     | `run`         |
| Bangla                    | `bn`          | Irish                     | `ga`          | Russian                   | `ru`          |
| Bashkir                   | `ba`          | Italian                   | `it`          | Samoan (Latin)            | `sm`          |
| Basque                    | `eu`          | Japanese                  | `ja`          | Serbian (Cyrillic)        | `sr-Cyrl`     |
| Bhojpuri                  | `bho`         | Kannada                   | `kn`          | Serbian (Latin)           | `sr-Latn`     |
| Bodo                      | `brx`         | Kashmiri                  | `ks`          | Sesotho                   | `st`          |
| Bosnian (Latin)           | `bs`          | Kazakh                    | `kk`          | Sesotho sa Leboa          | `nso`         |
| Bulgarian                 | `bg`          | Khmer                     | `km`          | Setswana                  | `tn`          |
| Cantonese (Traditional)   | `yue`         | Kinyarwanda               | `rw`          | Sindhi                    | `sd`          |
| Catalan                   | `ca`          | Klingon                   | `tlh-Latn`    | Sinhala                   | `si`          |
| Chinese (Literary)        | `lzh`         | Klingon (plqaD)           | `tlh-Piqd`    | Slovak                    | `sk`          |
| Chinese Simplified        | `zh-Hans`     | Konkani                   | `gom`         | Slovenian                 | `sl`          |
| Chinese Traditional       | `zh-Hant`     | Korean                    | `ko`          | Somali (Arabic)           | `so`          |
| chiShona                  | `sn`          | Kurdish (Central)         | `ku`          | Spanish                   | `es`          |
| Croatian                  | `hr`          | Kurdish (Northern)        | `kmr`         | Swahili (Latin)           | `sw`          |
| Czech                     | `cs`          | Kyrgyz (Cyrillic)         | `ky`          | Swedish                   | `sv`          |
| Danish                    | `da`          | Lao                       | `lo`          | Tahitian                  | `ty`          |
| Dari                      | `prs`         | Latvian                   | `lv`          | Tamil                     | `ta`          |
| Divehi                    | `dv`          | Lithuanian                | `lt`          | Tatar (Latin)             | `tt`          |
| Dogri                     | `doi`         | Lingala                   | `ln`          | Telugu                    | `te`          |
| Dutch                     | `nl`          | Lower Sorbian             | `dsb`         | Thai                      | `th`          |
| English                   | `en`          | Luganda                   | `lug`         | Tibetan                   | `bo`          |
| Estonian                  | `et`          | Macedonian                | `mk`          | Tigrinya                  | `ti`          |
| Faroese                   | `fo`          | Maithili                  | `mai`         | Tongan                    | `to`          |
| Fijian                    | `fj`          | Malagasy                  | `mg`          | Turkish                   | `tr`          |
| Filipino                  | `fil`         | Malay (Latin)             | `ms`          | Turkmen (Latin)           | `tk`          |
| Finnish                   | `fi`          | Malayalam                 | `ml`          | Ukrainian                 | `uk`          |
| French                    | `fr`          | Maltese                   | `mt`          | Upper Sorbian             | `hsb`         |
| French (Canada)           | `fr-ca`       | Maori                     | `mi`          | Urdu                      | `ur`          |
| Galician                  | `gl`          | Marathi                   | `mr`          | Uyghur (Arabic)           | `ug`          |
| Georgian                  | `ka`          | Mongolian (Cyrillic)      | `mn-Cyrl`     | Uzbek (Latin)             | `uz`          |
| German                    | `de`          | Mongolian (Traditional)   | `mn-Mong`     | Vietnamese                | `vi`          |
| Greek                     | `el`          | Myanmar                   | `my`          | Welsh                     | `cy`          |
| Gujarati                  | `gu`          | Nepali                    | `ne`          | Xhosa                     | `xh`          |
| Haitian Creole            | `ht`          | Norwegian                 | `nb`          | Yoruba                    | `yo`          |
| Hausa                     | `ha`          | Nyanja                    | `nya`         | Yucatec Maya              | `yua`         |
| Hebrew                    | `he`          | Odia                      | `or`          | Zulu                      | `zu`          |
| Hindi                     | `hi`          | Pashto                    | `ps`          |                           |               |
| Hmong Daw (Latin)         | `mww`         | Persian                   | `fa`          |                           |               |

## Install <a name="install"></a>

To install this program it's very simple, you just have to open the command prompt, go to the directory where you want to install the program and follow the steps below. 

- First, download the project from GitHub

```bash
git clone https://github.com/Kiplay03/excel-azure-ai-translator.git
```

- Second, open the previously downloaded GitHub project

```bash
cd excel-azure-ai-translator
```

- Third, create an `.env` file in the project's root folder

```bash
type nul > .env
```

- Fourth, open it with your code editor

```bash
code .env
```

- Finally, put your Azure AI Translator API key and its associated region available at the [Azure Potal](https://portal.azure.com)

```bash
# Configure the connection identifiers for the Azure AI Translator API available at https://portal.azure.com
azureApiKey="YOUR_AZURE_API_KEY" # Specify the Azure AI Translator API key to be used
azureApiRegion="YOUR_AZURE_API_REGION" # Specify the region associated with the previously provided Azure AI Translator API key
```

## Format Excel file <a name="formatExcelFile"></a>

Formatting the Excel file is also relatively simple. All you have to do is choose a column from a worksheet and put in the first cell of that column a [supported language or language code](#supportedLanguages) as in the example below.

![ColumnWithLanguage](https://i.postimg.cc/SNc4c693/Column-With-Language.png)
![ColumnWithLanguageCode](https://i.postimg.cc/2SmDG1F6/Column-With-Language-Code.png)

## Start <a name="start"></a>
 
## Improve accuracy with Custom Translator <a name="improveAccuracyWithCustomTranslator"></a>

Custom Translator is a feature of the Microsoft Translator service that enables businesses, application developers, and language service providers to create custom neural machine translation (NMT) systems without any machine learning skills. This feature can allow companies or individuals with fairly accurate translations that depend on a specific business need or vocabulary to learn artificial intelligence to identify these linguistic specificities and to translate and function. However, using this feature may require specific pricing, changes to this program, and lengthen the time of API calls. 

For more information, visit [here](https://learn.microsoft.com/fr-fr/azure/ai-services/translator/custom-translator/overview).

## Contribute and support me

As a passionate creator, I continually strive to improve my projects and offer quality solutions. If you find my achievements useful and want to contribute to their development, you can support me in different ways.

Your suggestions for improvements and new ideas are valuable, so don't hesitate to offer suggestions to enrich existing projects or to add new features. In addition, by sharing my projects and talking about them around you, you help to make them known and accessible to a greater number of users.

If you're a developer, your help is welcome! You can propose changes or bug fixes by submitting pull requests. Your expertise and contributions play a vital role in the continuous improvement of projects.

In addition, I invite you to subscribe to my social networks to stay informed of the latest updates, new features and upcoming projects. Your support on social media is a valuable way to motivate me to continue my efforts as a creator.

Finally, if you appreciate my work and would like to provide financial support, you can make a donation. Each contribution, regardless of its size, is greatly appreciated and allows us to continue to develop quality projects.

I would like to express my gratitude to all those who support my work. Every contribution, big or small, helps to move projects forward and provide ever better solutions. Your commitment and support are essential to help me grow as a creator and deliver achievements that enrich the community.

Project created and developed by [Kiplay03](https://github.com/Kiplay03)