PL-lang

# POTRZEBA
Podczas migracji do Office 365 dowiedzieliśmy się że kopiowanie poczty przez protokół IMAP nie przekopiowywuje spotkań w kalendarzu, kontaktów i zadań. (Przenosiny z Exchange 2019 w OVH)

# INFO
Działanie skryptu i przebieg migracji. 
  1. Migracja odbyła się w weekend. 
  2. W poniedziałek gdy użytkownik przychodzi do pracy i uruchamia Outlooka, ten mu pokazuje komunikat o błędnym haśle. 
  3. Użytkownik zamyka ten komunikat i na urchomionym Outlooku uruchamia skrypt.
  4. Skrypt tworzy plik .pst do którego kopiuje zawartość domyślnego katalogu z kalendarzem, kontaktami i zadaniami.
  5. Skrypt zamyka Outlooka i tworzy nowy profil. Po czym uruchamia Outlooka na nowym profilu.
  6. Użytkownik loguje się na swoje konto w M365.
  7. Użytkownik zaznacza w okienku skryptu że zalogował się już w M365.
  8. Skrypt podpina wczesniej stworzonego .pst i zaczyna przekopiowywanie danych z powrotem na główne konto.
  
Finalnie użytkownik posiada nowe konto w M365 z starą pocztą, kopią wydarzeń w kalendarzu, starymi kontaktami i zadaniami.
