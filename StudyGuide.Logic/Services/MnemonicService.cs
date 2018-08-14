using StudyGuide.Framework.Core.Boundaries;
using StudyGuide.Framework.Core.Models;
using StudyGuide.Logic.Boundaries;
using StudyGuide.Logic.Entities;
using StudyGuide.Logic.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace StudyGuide.Logic.Services
{
    public class MnemonicService : IMnemonicService
    {
        private readonly IBaseRepository<SentenceLibrary> _sentenceLibRepository;

        public MnemonicService(IBaseRepository<SentenceLibrary> sentenceLibRepository)
        {
            _sentenceLibRepository = sentenceLibRepository;
        }

        public string GetMnemonic(string abbreviation)
        {
            var abbrivArray = abbreviation.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            var count = abbrivArray.Count();
            switch (count)
            {
                case 2:
                    return Get2CharCombination(abbrivArray);

                case 3:
                    return Get3CharCombination(abbrivArray);

                case 4:
                    return Get4CharCombination(abbrivArray);

                case 5:
                    return Get5CharCombination(abbrivArray);

                case 6:
                    return Get6CharCombination(abbrivArray);

                case 7:
                    return Get7CharCombination(abbrivArray);
            }

            return "combination not found";
        }

        private string Get2CharCombination(string[] abrivArray)

        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Adverb,PartOfSpeech.Verb
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);

                if (mneme1 != null || mneme2 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get3CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb,PartOfSpeech.Noun
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Verb
            };

            var thirdComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb,PartOfSpeech.Adverb
            };

            var forthComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Adverb,PartOfSpeech.Noun,PartOfSpeech.Verb
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb,thirdComb,forthComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);

                if (mneme1 != null || mneme2 != null || mneme3 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get4CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Verb,PartOfSpeech.Adverb
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Verb,PartOfSpeech.Noun
            };

            var thirdComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb,PartOfSpeech.Preposition,PartOfSpeech.Noun
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb,thirdComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);
                var mneme4 = GetWord(theLetters.ForthLetter, comb[3]);
                if (mneme1 != null || mneme2 != null || mneme3 != null || mneme4 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3} {mneme4}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get5CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Verb,PartOfSpeech.Preposition,PartOfSpeech.Noun
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Verb,PartOfSpeech.Adverb,PartOfSpeech.Adverb
            };

            var thirdComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Adjective
            };
            var forthComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Preposition,PartOfSpeech.Noun
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb,thirdComb,forthComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);
                var mneme4 = GetWord(theLetters.ForthLetter, comb[3]);
                var mneme5 = GetWord(theLetters.ForthLetter, comb[4]);
                if (mneme1 != null || mneme2 != null || mneme3 != null || mneme4 != null || mneme5 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3} {mneme4} {mneme5}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get6CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Noun
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Verb,PartOfSpeech.Verb,PartOfSpeech.Noun,PartOfSpeech.Verb
            };

            var thirdComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Preposition,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Verb
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb,thirdComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);
                var mneme4 = GetWord(theLetters.ForthLetter, comb[3]);
                var mneme5 = GetWord(theLetters.ForthLetter, comb[4]);
                var mneme6 = GetWord(theLetters.ForthLetter, comb[5]);
                if (mneme1 != null || mneme2 != null || mneme3 != null || mneme4 != null || mneme5 != null || mneme6 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3} {mneme4} {mneme5} {mneme6}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get7CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Adverb,PartOfSpeech.Verb,PartOfSpeech.Adverb,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Noun,PartOfSpeech.Adverb
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Adverb,PartOfSpeech.Verb,PartOfSpeech.Adverb,PartOfSpeech.Preposition,PartOfSpeech.Adjective,PartOfSpeech.Adjective,PartOfSpeech.Noun
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);
                var mneme4 = GetWord(theLetters.ForthLetter, comb[3]);
                var mneme5 = GetWord(theLetters.ForthLetter, comb[4]);
                var mneme6 = GetWord(theLetters.ForthLetter, comb[5]);
                var mneme7 = GetWord(theLetters.ForthLetter, comb[6]);
                if (mneme1 != null || mneme2 != null || mneme3 != null || mneme4 != null || mneme5 != null || mneme6 != null || mneme7 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3} {mneme4} {mneme5} {mneme6} {mneme7}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get8CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Noun
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Verb,PartOfSpeech.Verb,PartOfSpeech.Noun,PartOfSpeech.Verb
            };

            var thirdComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Preposition,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Verb
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb,thirdComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);
                var mneme4 = GetWord(theLetters.ForthLetter, comb[3]);
                var mneme5 = GetWord(theLetters.ForthLetter, comb[4]);
                var mneme6 = GetWord(theLetters.ForthLetter, comb[5]);
                if (mneme1 != null || mneme2 != null || mneme3 != null || mneme4 != null || mneme5 != null || mneme6 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3} {mneme4} {mneme5} {mneme6}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private string Get9CharCombination(string[] abrivArray)
        {
            var fstComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Noun,PartOfSpeech.Adverb,PartOfSpeech.Noun
            };

            var secondComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Verb,PartOfSpeech.Verb,PartOfSpeech.Noun,PartOfSpeech.Verb
            };

            var thirdComb = new List<PartOfSpeech>
            {
                PartOfSpeech.Noun,PartOfSpeech.Adjective,PartOfSpeech.Preposition,PartOfSpeech.Verb,PartOfSpeech.Adjective,PartOfSpeech.Verb
            };

            var arrangeCombs = new List<List<PartOfSpeech>>
            {
                fstComb,secondComb,thirdComb
            };
            var theLetters = GetLetters(abrivArray);
            var allCombs = new List<string>();

            foreach (var comb in arrangeCombs)
            {
                var mneme1 = GetWord(theLetters.FirstLetter, comb[0]);
                var mneme2 = GetWord(theLetters.SecondLetter, comb[1]);
                var mneme3 = GetWord(theLetters.ThirdLetter, comb[2]);
                var mneme4 = GetWord(theLetters.ForthLetter, comb[3]);
                var mneme5 = GetWord(theLetters.ForthLetter, comb[4]);
                var mneme6 = GetWord(theLetters.ForthLetter, comb[5]);
                if (mneme1 != null || mneme2 != null || mneme3 != null || mneme4 != null || mneme5 != null || mneme6 != null)
                {
                    allCombs.Add($"{mneme1} {mneme2} {mneme3} {mneme4} {mneme5} {mneme6}");
                }
            }

            if (allCombs.Count > 0)
            {
                return allCombs.Shuffle().FirstOrDefault();
            }
            return "Sorry buddy!";
        }

        private LetterEntity GetLetters(string[] abrivArray)
        {
            return new LetterEntity
            {
                FirstLetter = abrivArray[0],
                SecondLetter = abrivArray[1],
                ThirdLetter = abrivArray.ElementAtOrDefault(2),
                ForthLetter = abrivArray.ElementAtOrDefault(3),
                FifthLetter = abrivArray.ElementAtOrDefault(4),
                SixthLetter = abrivArray.ElementAtOrDefault(5),
                SeventhLetter = abrivArray.ElementAtOrDefault(6),
                EigthLetter = abrivArray.ElementAtOrDefault(7),
                NinthLetter = abrivArray.ElementAtOrDefault(8)
            };
        }

        private string GetWord(string letter, PartOfSpeech part)
        {
            if (PartOfSpeech.Adverb == part)
            {
                var mneme = _sentenceLibRepository.Where(s => s.AdverbsPart.StartsWith(letter)).ToList()
                                                                                           .Shuffle()
                                                                                           .Take(1)
                                                                                            .FirstOrDefault()
                                                                                            ?.AdverbsPart;
                return mneme;
            }
            else if (PartOfSpeech.Adjective == part)
            {
                var mneme = _sentenceLibRepository.Where(s => s.AdjectitivesPart.StartsWith(letter)).ToList()
                                                                                           .Shuffle()
                                                                                           .Take(1)
                                                                                            .FirstOrDefault()
                                                                                            ?.AdjectitivesPart;
                return mneme;
            }
            else if (PartOfSpeech.Noun == part)
            {
                var mneme = _sentenceLibRepository.Where(s => s.NounsPart.StartsWith(letter)).ToList()
                                                                                           .Shuffle()
                                                                                           .Take(1)
                                                                                            .FirstOrDefault()
                                                                                            ?.NounsPart;
                return mneme;
            }
            else if (PartOfSpeech.Preposition == part)
            {
                var mneme = _sentenceLibRepository.Where(s => s.PrepositionPart.StartsWith(letter)).ToList()
                                                                                           .Shuffle()
                                                                                           .Take(1)
                                                                                            .FirstOrDefault()
                                                                                            ?.PrepositionPart;
                return mneme;
            }
            else if (PartOfSpeech.Verb == part)
            {
                var mneme = _sentenceLibRepository.Where(s => s.VerbsPart.StartsWith(letter)).ToList()
                                                                                           .Shuffle()
                                                                                           .Take(1)
                                                                                            .FirstOrDefault()
                                                                                            ?.VerbsPart;
                return mneme;
            }
            else if (PartOfSpeech.Pronoun == part)
            {
                var mneme = _sentenceLibRepository.Where(s => s.PronounsPart.StartsWith(letter)).ToList()
                                                                                           .Shuffle()
                                                                                           .Take(1)
                                                                                            .FirstOrDefault()
                                                                                            ?.PronounsPart;
                return mneme;
            }

            return null;
        }
    }

    public enum PartOfSpeech
    {
        Verb, Noun, Adverb, Adjective, Preposition, Pronoun
    }
}