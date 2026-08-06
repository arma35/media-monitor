Дополнительные CA для проверки HTTPS (вшиваются в media-monitor.exe).

russian_trusted_root_ca.pem / russian_trusted_sub_ca.pem
  НУЦ Минцифры (Russian Trusted Root / Sub CA)
  Источник: https://gu-st.ru/content/lending/

globalsign_gcc_r6_alphassl_ca_2025.pem
  Промежуточный GlobalSign (часто на *.gov.ru, пока нет в старом certifi)
  Источник AIA: http://secure.globalsign.com/cacert/gsgccr6alphasslca2025.crt

При проверке SSL программа берёт стандартный набор certifi + эти файлы.
