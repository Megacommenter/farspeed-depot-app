import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { storageGet, storageSet, storageSetGuarded } from "./storage";

const FARSPEED_LOGO_DATA_URI = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAQDAwMDAgQDAwMEBAQFBgoGBgUFBgwICQcKDgwPDg4MDQ0PERYTDxAVEQ0NExoTFRcYGRkZDxIbHRsYHRYYGRj/2wBDAQQEBAYFBgsGBgsYEA0QGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBj/wAARCADuAUADASIAAhEBAxEB/8QAHQABAAICAwEBAAAAAAAAAAAAAAcIAgYBAwUECf/EAFEQAAEDAwIDBAUGCQkGBAcAAAEAAgMEBREGBxIhMQhBUWETInGBkRQVMqGy0SM1NkJScnSSsRYXJENVVoKTwTNTYmOUwiU3ouFEZHODhPDx/8QAGwEBAAMAAwEAAAAAAAAAAAAAAAQFBgEDBwL/xAA/EQABAwMBBAcFBgUEAgMAAAABAAIDBAURIQYSMUETUWFxgZGhFCIyscEHFRbR4fAjQlJUsjM0U2Il8XKCkv/aAAwDAQACEQMRAD8Av8iIiIiIiIiIiIiIiIiIiIvPusdVPZKuC3ztp6uSF7YJSMhjy0hrj7DhegsXRtc0tPQrkHBBXBGRgr89rlXa70ZqusoK+83egusEhExbVPBJ68QOcEHqD35WyWrfXdSysY46hkrYDya2vgbM12OuHYBPxVlN5NpqHcLT4qaJsdPfqVh+SznkJR19E8+B7j3H3quW1+u37eawm07rChEtlmm9BXUlXEH/ACWQHHpA1wOMfnAdRzXslDc6S7290gpmvlYPeZgA97Tg/vRZGanmpZ9wSENdzW+WLtWVrOFmpdLQzDvmt0xYfbwvyPrUqad35201FwRNvvzbUO5CG4sMPPwDj6p+K4u+yW1eqacVkdigpTO3jZU2yQw8QPMOAHqn4KKdS9leuZ6SXSmooqkd1NcmcDvZxt5fELK42auBwQ6B3mPr9FaA3GDXR49VZ+CpgqaZk9NPHLE8ZbJG4OaR5ELtByqKzW7dzaWr9OGXmzxNd/tYXelpX+3GWY9oCkrRnamq4jHS62tDKiLoa+3DhcPN0R5H/Cfco1XsVVNZ01C8TM/6nXy/VdkV4jzuzAtParRotd0zrbTWsKAVmnLxTV0ePWYw4kj8nNPMe8L32uJznCyEsb4nFkgII5FWrXh4y05CzRY5OeqyXwvpERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERYuGe5QB2g9pfny3y640/Tf8AidLHmtgjHOoib+eP+No+I9isCsHtBY7i5jHep9suM1uqW1MJ1HqOYK6KmnbURmN/BVg7OW6JjnZt9farMT+drlefonqYc+B6t948FZ1uC0EEY7lTTfLbebQOtY9SWFr4LRXTelhfHy+ST54i0HuH5zT7R3KxGzu4rNwNvYaupewXWjIp66McvXxyeB4OHP25HctNtTb4Z4mXmi/05PiH9Lv159veq23VD2PNJN8TeHcpAlijnidFIxsjHDDmOGQR5hRbrXs/aH1YH1NJSGx3B3P5RQgBjj/xR/RPuwVLIHeuVlKOuqKJ/SUzy09hVpLDHM3EjcqkOptntyduLibzbfT1NPB67blaXuD4wO97R6w+sea27QHaYulA6K3a6p3XCn+iLjTtAmZ+u0cn+0YPtVq/Rg9SVFG4Ww2kdZelrqCMWW7vy75VTRj0ch/5kfIH2jB9q2EW1FJc2iC9xA9T28R3/p5KofbZac79G7wKkTT2pbFqizR3Sw3Onr6V/wCfC7PCfBw6tPkV63GD0yqLXSw7k7I6rjrmyzUWXYirqVxfTVI/RdnkfNruasFtlv8A2HWAhtWofQ2e9nkA5+IKg+LHHoT+ifcSoN12Vkgi9roH9LCeY1I7x++0Bd9Lcw89HON16mkHIyixjOYwcYWSyYVqiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi4J5rlcEZRFw5wGeeF8UFzt9VIWU1wpZnNPCWxTNcQfDAK+fUtqqL1pG52mlrXUc9XSS08dQ3rE5zSA73ZX5pao0vqXb7WFRZbzFNQ18JyJYnuaJmd0jHDHE09c/Hmqu5V7qMNduZB55W52M2Ph2kMkZqRHI3GBjORzPEcF+oDXHHPkgcCeq/Nywbg7wWy1mtsGpdTGhgdwOe1z6iJjsZw7iDgPet/092t9xbWWQ3uitN8Y36Rcw08p/xM5Z9yjxX6B2C9pbnyV3W/ZJdYifZZGS44gHB8jpnxV5gVkq7aV7Xeg7rI2n1Jbbhp+V3L0rh8ohz+s0ZHvaps0/rHTGqaEVWnL7b7nGRn+jTteR7R1HvCtIKuGcZjcCsJc9nrlazitgcztI08xp6r3VwfonKwEmc8sLknI6dVIVNla9rXStv1noiu07cGj0VSzDJMZMUg5sePMHB+Kp5oHUd12d3pfTXhrooGTfIrnD3GPPKQezk4HwyryFmR1wq19qHQvFDSa8oIA4sxSXAAdW/wBXIfYfVJ8wtnshXsL32up1im07jy8/nhU91gcGipj+JvyVkKeaKeJksMgkje0Oa4HIcDzBC71CPZw13/KHb92nK6bjuFmAiBcfWfTn6B/w82+4KbA7KzNxoJKCpfTScWnHf1HyVlTzCeMSN5rJYPABzhZrXtZawsuiNMT3y+1Iip4xhjBzfM/uYwd7j/7nkoscT5XCOMZJ0AHFdjnBo3ncF8e4F70nY9CVtXrGOnmtrmlhppWB5ncRyY1p6uPd4dchUSFun1Lqmqp9K6fqTHIXzRW+AundFEOeC49cD/8Acrbr9fdZ75bnQU1LTukc9xbR0LHH0VLHnm5x+049eg7grWbZ7ZWfbnTXyOjDam4zAGsrnNw6R36I/RYO4e8816VTSs2Spf4rt+okx7mfdaO3t/Y0WdkY66S4aMMHPmoE2m3+rtNmHTmtJp6q0g+jhrXZdNSY/Nf3uZ3eI8+itdb6+kuVBFX0FXFU00zA+OWJ4c17T3ghVj7S1g0Hb6uC50UgpNT1TuKSlpmgtnj75JB+afAjm7wPVR/tVu3d9u7syCd81ZYZXf0iiDv9mT1kiz0d4jvXxXbPRXykF0tzCx54sPA9e7+mh7CkNwdRS+zzu3h19Xerzd3VZrybFfrVqKwU13stZFV0VQzijljPxBHcR3juXq55LzZ7XRuLXjBHJaNpBGQuUQIuFyiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIuHfRKj7dPayw7oaQdbLnGKeuhBdQ17G5fTPI+th5Zb/qpCRdckTZGljxkFSaOsmo52VNO4te05BC/OjT1+1t2fd356WtpHh0ThHXW95Poq2DPJzT3+LXdx5eKuLRaY2m3e0jSakGmLRX01ZHxCf0LY5mH85rnMwQ4HkRlfNvptDR7oaPLqNkcOoaFpfQ1B5cfeYXH9F31HB8VV3Y3dK47T7iy6d1CJYLLVVHyeupphg0cwPD6XHdjo4d459yzzB93y9BNrE7hnkV7HUv8AxhbzdbcejroR74aSN4DmMeniOpTLqnsfaSrw+fSV7r7NMeYhqP6TD7OeHD4lQnqXYDd3b6qN3t1FLXRQnibX2KZxkbjvLBh4+BV/Kd7JaZskTmvY4Za5pyHA9CD3rnB4COQVhNZqaU7zRunrCydr+0u9UQEc7xMzmHjOnfofPKoxovtSbh6WmbQakZFqGljPC9tWPRVTB3jjA5n9YKzO3u/W3+4RipaO5fNl0dyNuuJEcjv1HZ4X+458l7Os9o9A69gc3UWn6aWpI9WthHoqhnmHt5n35Cq3uJ2U9V6dMlx0TUHUFAz1xTuxHVxgeA6Px5YPkom7X0WoPSN9Vetl2T2m917fY6g8CPgJ+X+Perv8Tc9QvKv9modQ6dr7LcIxJS1sLoZQRnkRjPu6+5Uc2/7Q24G3dwZZtSNqL1bIHejkoa8llTT+TXuGf8Ls+5W/2/3V0XuPavTaburJKhjeKahl9Soh/WYe7zGQrGhusc5G4d145cDkLIbS7EXGxjfmbvwng9urcdvV46dRKqfoW61202/zILk50cdNVOt1d3B0TnAcfs+i9XlY4OYC1wc0jII5gqqXah0k2j1TbdYU0X4K4MNLVEd8rB6rifEt5f4VM+x+rDqnZm1zTSmSsogaGoJ6lzOQJ9reEr0rahouNDT3hg1I3X9//vPovNbbmnnkpHcM5HcpKP0eSgnfXajWO4Oo7TXafqqaSlghMElNUymMROLsmQcsHIwD38lO6LJW24S2+dtTBjeHWM8Va1FOydnRv4LQNtNsrNtvpttFSBtRcJgHVle9uHSuHcPBg7h7+q8vdvdq27d2b5LSGOqvtQ3+j0pPKMH+sk8G+A7ypFvMldFp2ultcTJa5lO91PG/6LpA08IPlnCprpPazXW5m4c9ZqyG5UMIl47jX1kTmPce+OMHqe4Y5NCvbNTQXGeWvus3us1dni48h3d3cFBrHup2Ngpm6nTuXxaD0FqbebXVRdbrWVBozKH3C5yjJJ/3cfcXY5YHJoU/a+2H0xftEU9Fp6mgtFxt0JZSVAGGyNGTwSnqQTz4uoPPpyUmWKxWnS2nqe02aljo6KmZhjBywO9xPeT1JKrXvfvg67mp0do6qcLfkx1lfGcGoPfHGf0PE9/sVrFdLlfbiwUH8OOPh1NHWeRz1eCjPpoKOBxnG853mVoO2+5t72u1VLTH+mWt0xjraFrw5pIdwl8RHLiGOo5OCuxY73btQWKlvNqq2VNFUxiSKVh5OB/gR0I7iqn6Y7Oeor5tnU32snNBdZYxJb7fI3hD2jn+F/RLh9Ed3InqvO2k3PuW2Gq5rDqGOdlmlmMdXTytPHRyg4MgH2h3jn3Ky2htlJeukntzg6ePRwH83aOs/Ph1KNQVMtIWsqAQx3DsV1W9FyvnoqmCroY6qmmZLDK0PjkYctc0jIIPhhfQvLcY0WmHDREREXKIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi63NaXHuVUO1XtMHxO3MsNKOOMCO7RMH0m9GzY8R0d5YPcrYEOzyxhfLcKCmuFqqKCvp2VFNUxmKWF/Nr2uGCD7Qo1ZSNqojG5Xezd+nsdeyth5cR1t5j8u3VV07K+60t+06/b+9VPHcLbHx0Mjzzmp/0fMs5f4SPBWV4AWjqvzs1hZL5sN2gw61yvxRTtrLfMek9O4n1D48uJh9nmr96V1LbtWaMtmo7XJ6SlroGzMI7sjm0+YOQfYoNpqXFroJfiZp4LV/aFZIIZ47vQD+BUDeGOTuJHjx8xyXsEAEexYPa0sI+srJ8gAHIqqHaF7QXoJavQuhK7EgzFcbnC76HjFE79Luc4dOg59J1XVx00fSSH9e5ZOwbP1d9q20lKMnmeTR1lax2ote6I1Bf47DYrTQ1l2on4rL4wYLSP6lrh9PzJyB0CiWPSW4mkNO2zcSnt1ztlFIeOlucOWuZz5Odjm1ru7PIqY+z72fnagkpdca3oi21NIlobbMCDVnqJJAf6vvA/O9nWzO4WsdMaC2/qrlqowuoywwsouEONS4jlExh5HPwA5rPNoHVm9V1B3Ors7SvYJdrqfZ4w7PWiP2ndOH513idC1vEfQcOsqsTd96DcvaOu0Xr4Q0V+jjE9Bc2DhgqJo/WaHj+re4Zb4HPcve7LmpTR61uWl5X/AIG404qImnp6SPrj2tJ/dVc/mu4661vXjSOlnsMrpKplst7XSNp4hzwM9w/jyHcF6e2mtZdFbmWW71he2Gjq2+myDxMjJ4HgjrjBPLyWs2Z2tjbQTWi4/A/Vj+QcOGe8ga+apNuPsqbLILnYhiRoBfDkEgH+nHPjpwONF+mA6IvjoLnQ3O3wV1vqY6mlnYJIp4iHMkaRkEEdQvrDgV8heQuaWnB4o5oc3hIyCuOBvn8VkiLhVe7RG690hvNTt5YzJSQxsb8vqmnD5eIAiNuOjcHme/OOi+rY7Y0wGn1nrOj/AApxJQW6Vv0O8SSDx8G93U92JxumgNI3nVlPqS56foqq6U+BHUyNy4Y6ZHR2O7IOFsTGuBJIx71qZNo2w25tBQs3Mj33cyew9X/rvrG28vqDPM7e6h1IGDgz/BQD2g9qW3y2v1tp+k/8TpGf02GNvOpiA+njve36xnwVgTlYSMLmYwD5FUtsuU1tqW1MPEcuscx4qZU07aiMxu5qsHZ03VdBJHt/fKvEbyTa5pHdCeZhJPd3t+HgrQNdk9VTbfvbR2htYR6lsTHwWm4Sl4EXL5JUZ4uFp7gfpN8CCO5T/stuHHr7QcctZI354ocU9azvccerLjwcPrytPtPboaiFt5oR/Df8Q/pd+vz71W22oex5pJviHA9YUmouC4LkHIysSrlERERERERERERERERERERERERERERERERERERFi/6KyWLugRFAPao0A3Uu1g1PRQcVxsRMxLR6z6c49I3zxyd7itQ7IGu3TUVz2+rp8mnzX0AcfzHHErB7CWux5lWir6KC426ehq4hLTzxuilYejmuGCPgV+etEavZTtRtifI5sVpufonk8vSUj+8//beD7QqC4A0tWyqHA6Fet7Iu+/7BV2GT44x0kf1A8f8AIr9BbnTfLbVVUInkg9PC+H0rOTmcQI4h5jOVWjbzsnNsmuPnPW91obxQUr+Olo4GuAnIPJ03EOg68Izk9TjraKmfHNEHxuD2OAc1w5gg8wu1wAYcD4BW01JDO5r5BnHBefWvaGvtUE1PSSbglGHaDPgeI58Fp+udcac210bLf77UCKFg4IKaPHHUP/Njjb4/UBzKo1ebxr3tC7vxQwwGWeQltLRsJ9BQQZ5uce4fpO6k8h3BWR7QeymrdztRWW5abudE2KlhdTy0tZI5jY8uyZG4ByegI68gpA2p2osW1mlhQ24Cor5wHVte9uHzv8vBg7m93tVbVU09ZN0TvdiHqtps9erTs3bPb4cS10mQARpH++PWeGgym1W1Fg2v0kLfQD5RXzgOrrg4YfUOHcP0WDub/qo7347PlFrGmqNVaMpY6bUTB6SembhrLgAOfkJPA9/Q+KsO3n1WRA8ArGSiifF0BHurH0e01ypLj96NlJlJySf5uwjq7OXLCoXsjvfc9rb67S+qW1LtPmYslhkB9JbpM+s5oPPh/Sb7x53qttdSXK2w3CgqY6mlnY2SKaN3E17SMgg+BVfe0PsYzV1rn1ppWiazUFMziqKaMAfL4x5f7wDp4jl4KLuzlvdJpC7RaF1VVEWOok4KSomOPkMrj9E56RuP7p8icVNNPJQS+zTnLTwK9BvtrpdrKF19tDN2dv8Aqxjn/wBgOf17wrvIuljw5wwQu5aBeQIiIiIiIiLX9Y6Yt2sdI12nrozMFVHwhw6xv6tePMHBVM9LXm+bKb2SRXFj8U0ppa+EZ4ZoCc8bfHlh4V6z1Vc+07oT5ZZqfXlBD+HosU9cGj6URPqvP6pOPY7yWx2RuMbZHW2q1im07ncj48O/CqLrTktE8fxM18FYOiq4K6ghrKWZs0EzGyRyN6OaRkEe5fUOigLs067N40dNo6tkzWWn1qfiPN9OTyH+E8vYWqfG/QHsWdulvkt1U+ml4tPmOR8lYUtQKiISN5rlERQF3oiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIhGUREWPD5ql3bD0u2j19ZtVQswy5UzqWcgf1kXNp97XfUrpqCe1Zp9t27P89xaz8LaqyGqBPc0n0bh8HhV12h6WleOY18ls/s+uJoL9Tvz7rjuH/7afPC2rYHU0mqtgNP19RJ6SpghNFM4nJLojwc/MgNKkxVe7Gl7dNo/Umn3yZ+S1cdTG0no2RmHfWwfFWgDgeWea7LdN01Mx/Z8lB2xt4t96qqZo0DiR3O1HoVj6MDouQzlzOVkimrNLjHPquUREXW6LiGOIj3Knvaj2bbaa1+4+m6MNoqh2LrBG31YpHHAmAHQOPJ3ng96uMvgu1vpLra6m219NHU0tTE6GWGQZD2OGCCPNRK2kbVRGN3gtBsxtDPYK9lZDw4OHW3mPy7VAXZi3dOqtPDRF/quK9WyL+jyyO9argHL3uZyB8Rg+KsRxuz3L87dcadvmxO/Uc1pkkYylnFda53/ANbASfUJ7+9jv/dXw0Pq22630FbtT2t4MFZEHlmcmJ/RzD5g5ChWqqc4Gnm+NnqtLt/YYIHx3i3f7eoGR/1cdSOzPHHLUclsrTluVyuG/RC5VwvOkRERFwRlfHdbdSXSy1Vtr4hNS1UToZWEfSa4YK+1Yv5sXLXFpy3iuCAeKonQTV+zPaADZXSFttq/RS/8+lf3+eWEH2hXnp6mKqo4ammkbJDKwSMe08nNIyCPcq19qbSIDLVrSmh55+QVZHnl0ZP/AKh7wt57O+rHak2jgt9RLx1dnk+Rvz1MeMxn4cv8K3e0X/lLXT3ZnxD3H9/X5/NUdvJpql9I7gdQphBXKwYSTzWawQV6iIi5RERERERERERERERERERERERERERERERERFpW7Np+fNkNU2zHOW2zFvLo5reIH4tW6rzb5A2p01cadw9WSllYfewhdczd5hb1gqVQzGCpjlbxa4HyOVTHseXJ1PvDdbdxYbWWovx5xyNI+0VdzkHBUB7MM76PtMWqnYTwywVUDvYIyf8AtCv+0AgclUWE5pcHkSvQvtahEd+L2/zMafmPos0RFdrzNERERFg4HiBys1wQD1RFDPaK23ZrzamprKKn471Z2urKRwHrSNA/CRf4mjl5gKFeyRuCbZqyr0DXz4pLkDVUIceTZ2j1mj9Zoz7Wq5sjGGIgtBHgea/PfeDTNbtF2iXVtiPyeL07LvbXM6NBfks9zg5uPAqhujPZpmVrO49y9Z2FnberbU7N1B1IL488nDiB44PdlfoUw5jafJZLwdJ6jpNV6ItOpKJ4MFfSsqAAehcObfaDke5e6Dk9VehwcMjgV5TLE6J5jeMEEg94XKIi5XwiHoiIi1PcbTTNXbZ3iwOaDJUU7jCfCRvrMP7wCrB2btRyWLeE2SpPo4rtC6nc13LEzMubnz5Ob71clzGkZ4RkKjW4tJNt72k6ysom+jbT18dzgA5eo8h5Hs+kFu9kD7ZTVVqf/O3eHeP2FR3UdFJHUjkcHuV443Ek8l2L5qCpirbfBWwO4op42ysPiHAEfxX0rCYIOCrwHOoRERERERERERERERERERERERERERERERERERERfJXYFtqA44Hon5PuK+teTqKobS6RutU4gCKjmeSe7DCV8vOAuyFu89oHMhUO7N44u1PZQzOAas+70T1+gjegVBuytTuqu0jQVHDn0NFUzHyyzh/7lflv0QqSwf7YntP0Xpv2uHF6Y3qjb9Vmi4BBXKvV5ciLjIz1C5RERMjxC4zzREcCWkBVv7XekBc9tbfq2njzUWeo9HKQOfoZcA59jg0+9WQJC1rXmn4tVbbXzTr2h3y6jlhbnueWktP7wCi1sHTwPj6wrvZu6G13SCsB+Fwz3HR3oSoR7IWrDdNt7jpOokJmtFSJYgTn8DLk4HscHfFWQY45wVQjsyaph0jvqaW71kVFSV9LLSTvneGMZI312lxPIc2uHvVt6/enbC2yFs+s6F7gOlPxTZ97QQurZ+KespmiJhcW6HAJV/8AafQxW2+yuyA2XDx48fUFSHk+CZUQTdpPa6Fxa243Gcj/AHdC/B95wusdpjbAnBqLsweJoXH+C0wsFzIz7O//APJXnJr6cfzjzUxF2Dz/AIrkOBOFFlH2gtqqsDi1E+lPhU0sjMe3kVtVp3C0Re3N+atWWioc7mGCpa1/7pwVFmtlZBrLC5veCPou1lTE/wCFwPiFtSqr2rbK2LU1gvzWYFTBJRyHpksdxN+px+CtPG9r2BzXBwPQjooU7TtrbXbNsrwwOfQV8UmfBr8sP8R8FbbJVJp7tC7kTunx0UW6xdLSvHitr2QvJvexWn6iR5fLDAaWQk55xuLf4AKQ1APZTunynb282ouyaS4ekaCeeJGD/VpU/ZHiol/pvZrlPF1OPrr9V3UEvS07H9iIiKnUtERERERERERERERERERERERERERERERERFH29N1Fl2B1bXcfA/5vkhYfBz/UH1uCkFV57XV+bbtlqazNkLZbpcI28IPVkYL3e7PD9Si10vRU739hV7sxQmuu1NTgZ3ntz3A5PoFFvY5txm3QvtzMZ4KW2tiDsdDJIOXwYrolwaOZ8lSvYPc7Rm022F8u1+nfPdLlWNbT2+lAdM9kbMAnuY3ic7mT3LxNadpbcrWtc626c4rFSTHgjprc0y1Mg8DJjOf1QFS0NfBRUrWuOXHXA1Xpu1WyV12kv1RPC3chbhu+7QYAGcczrnhp2q6GpddaQ0dB6XUuorfbQejZ5Rxu9jfpH4KH7/2u9ubdxx2Whu96eOjo4hBGf8Tzn6lA+l+zfuvraoFzutOLRFMeJ1VeZXGZw8eDm8n24Uz6c7HujqPEmpdQXS7P5Esp+Gmj9nLLiPeu0VdfU6xRho6z+/oq51g2Ps+lxrHTyDi2Ph5jP+XgtLunbK1HIXizaKtlM3819XUPmPvDQ0fWtZqu1puvUOPyd1iph4R0fHj95xVpLXsHtFZmMFNoe3TOaeT6vindn2vJVfu0fT6etu4VtsFhstuoI6Kj9JK2lpmR8T5HZGeEc+TR18Vb2PZu43irFMajdzk5AzgDyUSs2z2PtsZfBai//wCTvzLlqQ7VG74IcbnanDPR1vYAV91L2ut0IXAz0+n6od/FTub9ly3f+drS+h9qrLpbTtktt6vUNG35TWT07XQQvcOIgnGZHDix4cuqjq0aJ1tu7qKSupLRDUcTuGWsdAynpoR4AtaBy8Bkq7p/s9um4+eerEcY4F4xntxnTKhO+0nZZ7hEbMCTx3SNPQfNb7aO2bWtc1t90LDI0dZKGsLXHz4Xj/VSZp/tR7VX6RsNbXVtkmd3XGDEY/xtyMe3C8HTnZI0VDbXO1VW1dwrHt6UTzTRRHy6lx8z8F4GpexxQvEkukNXTwP/ADae5xCRp8vSMwR8CspURXCleWxyNlaOYBGfMBWsc2w10ADmSUzj4gf5Kveuqa3Q763mKzVlLNQTXRz6eoieHxFkjwQQR3esfgrF2zspXSRmbprGjgBxgUlK6T63EKGq/s5bu2zUMNvi018rD5AGVtJM18Def0nEkFo7+YV/7bFNFbKaCofxSshY17h3uDQCfiCrLZO93O1snjhO4HOydBx7Mro+1mK13BtA+nmbMWMLSQRwGME4PPVQJB2UbA1o+Vauukp8Y4I2fxyu93ZU0vwnh1Tege4lkRH8FP8AwDxKy4QtOdrLuTn2g+n5LyD7qpf6B6qtdZ2T6XB+b9a1DD3NqKRp+sOWpXbsva6o2mS3XGzXIDmG8ToXn94Y+sK4PCFwWA9VLg23u8XGQOHaB9AD6r4fZ6V3BuFRd9u3n21lM4ZqO1Qs6vikdNTn244m49q+66786o1Jt1c9J6moLfXtq4uBtZEPRSRuBBDnAZaenkrsmJhBBHVaJq3Z3QOsGvluNjigq3f/ABdF+Akz4nAw73gq0g2uoql7XXGlG8CDvN46eR9VEktc0bSIJTjqPBQh2Ua8x6v1FbuI/hqOKcD9R5b/AN6tSBlvuUZ7Z7K2Tbe9191o7pW3CpqY/QNdUNa0Rx8XEQA3qcgc/JSeGgDHNZ3ae4QXG4vqab4Tjs4ABT7ZA+np2xycQuURFQqeiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIipH2wr3VVW7drsjo5GU1DbxLGXcmyPkceIt8ccLQruLx7lp2yXetp6u62mgrpqZ3FA+pgbI6I56tJHJQ6+kNVCYg7GVpdkr9HYri2vkj390HAzjUjGeapFtb2Z9W66ZDd7+ZNP2R+HNdLH/SZ2+LGH6I83fAq3mhtqNFbeUoj05Y4YpyAH1sv4Sok9rzzHsGAt1jYGudhdi66O2QUo90Zd181L2k24ul+cRNJux/0N0b49fj6LDgOSuA12F2IrBY9dbxhoz44VB9za+s1Jvlf5GU80s8le6mhga0l5DMMa0Drk4+tX6eARzXjs0tp2LUsmoY7JQtusjeF9aIW+lcPN3VaLZy+ts0z5zHvEtwNcYVfcKI1bAwHAzqq/7Zdm1pbDe9wGl7uT4rOx2APD0zh1/VHLx8FZCioaa30UdHRU0VNTxNDY4YmBjWAdwA5BfQ1oGCFkoF1vFVdJOkqX56hyHcP2V3UtHFTN3YwsC0rjgPhzXYirFJwun0ZLuYWbGkOOenis0RERERcoiIiIuCFyiIuAMHmuUREREREREREREREREREXXxta3JOEznmM4UQ7811dRWCzmjq56cuqXhxhkLCfU78KDG3y/ucGsvFycfAVEn3rPXC/sopzC5hPBZa6bUR0FQad0ZJGOrmro8/Fyc/EqmPzxqPGfnW6/50n3rn531Hj8aXb/Ok+9Q/wAVM/4j6KD+NI/+B3ormBwzgke9ZE4OFXXZS4Xiq3KdHWVtbNF8kkPDNI9zc5bg8+/qrFHor+31orIel3d3sK01ruAr6cThpbnkVwXgDmVxxtBznkVHe8WpH2DQT4KWoMVZXvEERYcOa3q9wPs5e9VyN+vmPx1ceXTFS/l9ar7lfY6KXot3JVVdtpYrdN0DmFxwDp2q6rTkZQnAWnba6kOo9uaGtmeH1UYMFRz58beWT7Rg+9bjnKuYZmyxtkbwIytBTzNnibKzg4ZWDnADmea4xy71Ee/VZW0Wn7O+irJ6Zzql4cYZCwkcHfhQY2+X554WXm5OPg2okJ/iqO4X9lHOYSwk9izt02njoKg07oy4jHBXQBxyOU4m93PKpozUOp6X1mXu7RZOc/KJBz95XuWbdTW1nq2PN4krog4F0FXh4cPDPUfFdEe1EBcBIxzVEi2zpy4CSNzR1q2I6dVx3EZ6LwNI6opdWaWp7xSM4A/LZIycmN46tK+TXesqfRemH3CVgmnkd6Onhzgvdj+A6laF1RGIunJ93GfBao1cQg9oLvcxnPYtoy0nKekYTjjafYVUO9a61bqStLqu61RD3fg6alcWMafABvM+/K8+Sg1NTsM8tHeYmt5+kcyUY88rNu2na4/woiQsk/bNrnEQwlwHPP6FXM4xywSQueIHKqRp7cXVmnamOSnus1VTgjipapxexw78E8x7irPaV1BSap0rS3qj9Vsw9aM9Y3Dk5p8wVa227xV2WtBDhyKurTfoLllrBuvHIr2S71gM9U5+JUP79VldQ2WzPo62enL55A4wyFhPqjrgqD23y/udwtvNycfAVEhP8VEuF/ZRzmEsJOnBQrntPHQVBpjGXEY4dqugC4dx+CF/Ln9apq3UGp6dmWXm7RjvPp5B/qvVtW5mtrVUMljv1TUsb1iqz6VpHnnn8CozdqocgSRkKGzbanJw+JwHgrajp349qcbWtyThahoHXdFrexOnij9BWQkNqacnPAT3g94PcVpu/VdXUVmsxo6yenLqiQOMMhZxYZ346q7nr42Upqm6jC0VTc44qQ1rfeaBlTDnPMZwnPxcqXNvt/c7hZd7k4+AqJPvWfzxqPH40uv+dJ96oBtXGeER9FmhttE7UQnzCubz8SuQ7uKpl876j/tS7f50n3qfNjqmvqdD1j66eoneK0gOnc5xA4G+PcrC3XxtbL0QjI7VZ2raNtwn6BsRbpnJUo58CnFnvPwUQ79VtbRafszqKsnpnOqXgmGQsJHD34Kgxt9vzyGtvFxLu4NqJCf4r5r9oGUcxhLCTp6rrue07KCoNOYySMcMc1dHn/xLjmeR4lTP531H/al2/wA6T71x88aj/tS6/wCdJ96h/ilnOI+ig/jSP/hd6K5zXc8ZKOcAOZ5qKtiqquqtJXF1fU1M0gq/VM7nOOOAdM9y+XfusraKy2V9FWz0xfUSBxhkLCfUHXCuTc2+xe2FumM4WgN2aKD28t0xnHope546lM4PMnCpg2+357gGXi5OPgKiQn4ZX0xar1XRSD0WoLrCe4fKHjPxVMNq48Z6I48Fnm7bQnXoTjwVxs5cDzxldqrNpfenUlprY4748XWhOMlwAlYPEOHX2H4qxFovVDfLNBdLZOyennbxMe3+B8CPBXVvusFcP4R1HIrRWu801xbmE6jkeKiftBfk/Zf2qT7CjzaAxDdu3mfg4PRTZ48Y+h5qQ+0F+T9l/apPsKCKWkq62qEFFTS1ExBxHEwucfHAHksheZTFdQ8DON3TwWEv85gvQlAzjdOOvRXPElszykoyfaxZekt/6dF8Wqn505qcdbDdf+mk+5c/yd1N/YV0/wCmk+5WX4kk505/fgrb8Wy/2p/fgriwMpCPSQCE5/OZj/RdjzgLQdnqWro9raanrqeanlE8p4JmlrgOLwPNbHq6+x6a0TX3mUjigiPowT9J55NHxIWliqGmATuGBjPctdBVNfTCoeN0YyR4ZVet49R/Pm4ctHE/ipraPk7AOhf1efjge5eDVaWqabbOg1aGvLKirkgcO5rRyafeQ76l4tPDUXW7Rwh/FUVUwbxOOMuc7mT8cqzV6sFkm2em0nTVtKRDSAQuErcmRoyD173D61hKWn+85J6h/Vp38vRea0lIbxJU1Mh5HHfy9Ao22J1B8h1XUafnkAirWekiB/3jR097c/BWIBySFSm03Gqs19pLnB6s1LK2UDzB5j39Fcu1XCnu1nprlSkOhqYmysPkRlXmzFV0lOYXcW/JabY6u6WndTOOrD6H8lE/aD/Jyy/tT/sKOtnmB+7lva5ocCyXIIyPoKRe0H+Tdl/an/YUD0hrGVQfQGoE4BwYM8eO/GOfRU14l6G6dJjOMHHgs/tBP7PeRNjO7unHgFc6opKKSnfHVU0DoiMESsGCPPKqVreC0024N1gsTmOoWzYj4DloOPWDfLOV5U9yuU3FFUV9ZIPzmyTvI94JW0aK23u2s3CamqaSnoWO4ZJC8Oe3y4B3+3C5rrgbtuwU8WDnsK4ud1dfNynp4fezx0Kk3s+tqRpe78YPyc1TfR56cQYA7/tXy9oSkqHUllrWtcaeN8kb/AOcGkfwKlbTWn6HTGnqez25hbDCObj1e48y4nvJK7b5Zbdf7PPbbpA2enlGHMPd4EHuK0/3Y824UZPvYwtj9zvdahQk+9jj28VVjb7U9v0nrSO6XKgNVAY3RktAL4ices0H2Y96sNa9zNEXhvBHqCmjJH+zqcxH/wBXJRbqDYi8U8z5dO10NZTnJEM59HI3yz0P1LRLpoLWVoY51dp6r9G3rJG30rR725Wcpp6+0tMZiyz98wspSz3WytMRhyzPUT6hTjUbP6FvdfUXOKoq8TyGQikqG+jBPXhABwFuOlNKWzSFmfbLVJUugfKZSJ3h5BPXHIKptn1BetO1wqbRcJ6ORhyWMceF2O5zeh94VnNudas1ppc1UkTYq6neIqmNp5ZxkOHkRz+KtrNXUdRKdyMMk+firrZ+5UFXOejiDJccufXqtM7Qv4isn7RJ9kKPNoGsfu7bWvAILJeR7/UKkPtCfiOyeHp5PsqDbdcq603Blfbat9LUsyGyxkAjPI/UqW8SiG6CQjIGM+SoL7M2nvQmeMhu6fRXQdT05yDBGR+jwAghVt3qtlptm4EPzXDDA+amEk8cOAA7iIBwOhIWtybha2mZwP1Tccd/BLg/UF51Ja7/AKiuThSUVbX1MrvWfwucSfEuP3rtud6ZXRdDDGckrtvN+iucHs8ER3sjXHDuwpD2EdMNf1zGl3ojREuGe8PGP9VsnaD/ABLZP2mT7C2Xa7b5+jbRNU172PuVWGmXg5iNo6MB7+pJPita7Qn4lsn7TJ9hWDqR9LZnRv48fMhWbqOSk2ffFLo7jjqyQo82mqqGi3So6i41EEFO2GYOfO4NaCW8skqxg1No0H8fWUnznj+9VJtVouF9ujbdaqV1TVPaXNiaRzA69VsI2r16T+TM+f1mfeqm0XKopodyKEuGTrgqlsd2qqSn6OGDfGSc4PkrLfyn0f8A27ZP8+P7169JLR1FIyeifDJDIOJskJBa4eII6qqn81Wvv7sVH7zPvVktD0FXa9vrPb66Ew1MNMxkkZ6tcOoWotlfPVPLZYtzHNbGz3OprZHNng3ABx1+oCjrtCfk9ZAOnyl/2FHG01VQ0e6NDUXGogp6dscvFJO4NYPUOMk8uqkjtC/k/ZP2p/2FB1ptNwvlzjttrpXVVVICWxNIBIHM9Vl7xI6O677Rkjd06+xZC/SuivXSMbvEbuB16DRW3/lPo/8At2yf58f3rE6n0f1+frJ7p4/vVbf5qtff3ZqP3mfesTtZrwczpqfl/wATPvVp9+Vv9sfVW/4juH9mfI/krW0ctLUUrKiikhlgkHEySEgtcPEEdVEXaE/ENj/aZPsBSDoOgq7Zt3Z7fXwmCpp6VrJIyebT4FR72hPxDY/2mT7AVld3F1ue4jBIHzCur44vtT3OGCQNOrUKPtnGtk3boWua1wMUpwRn81WUr7JarrSGmuFtpamFwxwyxA8vgq3bMf8Am9Qf/Sm+yrR/mKJsy1rqMhwzr+Srtjo2voCHDI3j8gqpbnaMi0bq1sNGXGgqmmWAO5lmDhzc94Hd7Qtv2E1DPFeq3TUspdBNH8pha4/Re04dj2gg+4r7e0MW40/04szcvLDVqOygJ3epOXL5PNn91VG4KS8hsWgyPUKiEQob+GQaAkDwIGVvXaC/J+y/tUn2FH+zQB3jtw/5U/2CpR3o0/e9R2a1w2O2zVr4ah75BFj1QWYz1Wm7XaK1XZN0KG43WxVVLSxxyh0sgGASzA7/ABUqvppXXdkgad3LdcaKZcqWd19ZK1h3ct1wccsqw7efUn4rLh8z8VixZrZjgvQl0lrWuJxzUGb96l4zQaXgeSB/S6jHwYD9Z+CnSR2I3HBOO4Kr+p9Ibgak1lcLxLpmuPp5SWA8PqsHJo6+ACotoJZBTdFE0ku00WZ2plmFH0UDS4u006lpttst4u7JPmm2Vdb6LHGaeMv4c9M48V9/8idY92l7uP8A8dysNtNpSp0vocNuFK6CvqpTNOx30m9zW/AfWpBwVV0mzDZIWvleQ48VT0OxzJIWSTSODjqQOSpLX2u42mqFNc6Goo5i3jEc7Cx2CeuD3KftitRGv0nUWKeXMtvfmMHqYncx7gchN6dFXLUcduuNkoJKusgcYZWR4yY3DIPPwI+taftrpnXOltf0ldU6dr46OXMFQ7AwGno48+44K6KSkltlxDWNJYdM45H9VFoaGe0XUNY0ujOmccj19xWzdoP8mrMf/m3/AGFHmzgB3dtwPMGOX7ClPeiw3rUOnbXFZrdNWSQ1DnyNixloLMArS9rtE6ss25lDcbnYqqmpY2SB0rwMDLcDvXbX00rrs2QNO7luvku65Us7r2yVrCW5brg45LZ939uWXK3P1NZKYNroG8VTDG3/AG7B1cAPzh9YUL6X1Nc9J3+K622UhzcCWIn1ZmZ5td/oe5XHc3ibwkKvO4u1F1h1W+u0rbZaqjqyXuhhA/AP7xz/ADTnPlzXbfLZJHIKukHvcwPmu/aOzSxyivoh73MDr5HT1U4ad1BQ6m09T3i3ycUMzM8Pex3e0+YKi7dfXOtdJ6uiprXWxQW+ohEkTzTtceIcnDJ9x96+Hau3a90jqA0Vw07Xi01jvwpIBEL8cn9encfcpS1fo+3az08bfXAxuaeOCoZ9KJ+ORHl4jvViZKiuocsyyTxGo/NW3SVVyt2Y8xyjHWNR9Coy213bqKy6VNv1ldWZn4XUtRKxsbGnvYSAAM8sEqZDcrb6L0hr6TgxnPpm9Piqz3rZ7WtpqHCmt7bpBk8MlIQSR5tPMLwG6L1g+Y07dM3Uv/REDv8A+KoprvXUjehniLiOB1/I5VFSXy5UMYp6iAvI56+uhyva3YuNkuW4009iMT4hE1sssP0JJBnJGOR5YBK3Xs+QziW+VAyKfETOXQu9Y/UD9a1rTuy+rLtOx1xhZaqXPrumIdIR5MH+qsDpjTVu0pp+O0WuItiZ6znuOXSOPVzj4lfVpoKiWsNbK3dHUvux2yrmrzXzs3BknHeOQ+qjLtCfiSyD/nyfZUdbSU1PV7sW6CqhjmicyXLJGhwPqHuKljerT161DabTDZbbNWyQzvdI2MD1QWgA81pe2OiNWWXc2guN0sNVTUsbZA+V4GBlpA71811NI+7NeGEty3XGnDVfFyo5n3xkoYS3Ldcacsqeo7FZWYcy0ULT4iBgP8F9rYIWN4WRtaPBowsmEGMELJbMMaOAW/axreAwseEBvJQr2hPxNZP2mT7Cmw9Col3o09e9R2u0xWO2T1roZ3ukEWPVBbgHmVW3mN0lHI1gySPqFU7Qxukt8rGDJOOHeFFm0VVTUe69FUVc8UMTYZgZJHBoBLfEqyg1Hp7AHz3bv+pZ96q9/Nlrwj8l60+5v3rj+bDXf91az91v3rK2yvq6GHomwE655j6LF2e5V1tg6BtM52pPAjj4K0f8o9Pf23bv+pZ96+mjuVvuBcaCup6kMOHehkD+HPjgqqn82Ou/7q1n7rfvUs7JaZvmnPnkXq1zUJmdF6MSADiwHZ6e0K9oLtU1EwjkhLQeev5LR22+VlVUNilpy1p56/UBfP2hPxBZP2l/2FHW0NTS0e69BU1c8cMTY5cvkcGtBLCBzKlnerTt81FZrVDZbbNWyRVDnyNix6oLcZOSoZG2OveHnpat94b96pbsydly6eOMuA3TwPUqC9xVDLv7VFG5wbunQHkFaP8AlHp/+27d/wBSz71idSaeyM3u3f8AUs+9Vf8A5sdd/wB1az91v3rg7Ya7/urWfut+9Tvv6s/tz6/krEbS3D+0Pr+StXQ3G3XBzzQVtPUln0/QyB/Dnxwon7Qn4hsf7TJ9gLv2T0zfdOPvHz1ap6EziL0fpABxY4s9D5hfVvZYL1f7PaYrNbpqx8E73yNjA9UFoAKm1sktVbXOczDiOHirOulmrbQ97oyHuHw89D+SijaOqpqLdaiqKudkETYpcvkcGger4lWErtf6Ot1OZKjUlvw0c2slD3H2BuSq3fzZa9I/JetI9jfvXdT7Va9nkDRp6WHzmexg/iqG21tbRxdDHATrnJBWWtFwuFup+gjpydc5IP5JuRrNutNWirp2OZQ0zPQ07X8iRnJcR3ZOOXgFvmw2m5xUVmqJ2FsZb8mpyfzueXuHlyA+K69M7CVklQyo1TXxsgByaSlcXF3kX93uU4W+3U1soIqGigZBTwtDI42DAaArC1WuokqTW1g14gfvhhWlls1VLVm4VwweIHPPLuwu9rRjyXBaM8hge1d6LWhbpYDuxhZoi5RYFucjnhdZaF3pgeCIsG9MYws0REXVI3LyU4Wk8x712pgIi6eEDngYRoAdyC7sDwTARcLDvXW5oLskZx5rvTARcrpLQW4AwsmN4WcK7EXGFxhdRYD3Z9q5wD3k+WV2JgJjK5XS1oz3/FZgcyfFZouUXU5ozl3NcOaO4Lu6ouMIsWDDAFkiLlEPQrpa3lzC7kRF1gAd2Ex5rsRcYRdePNY8PrZ8+8ruRAEWBGfBccI8l2ImEXXjzQtyuxEwi6g1oOQCjmjOT/FdqYC5RdYCEcx967MBMDwXGEWA5ADCzRFyi//Z";
const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');`;

const FONT_DISPLAY = "'Barlow Condensed','Noto Sans TC',sans-serif";
const FONT_BODY = "'Inter','Noto Sans TC',sans-serif";
const FONT_MONO = "'IBM Plex Mono','Noto Sans TC',monospace";

const LIGHT_COLORS = {
  bg: "#F3F1EA",
  surface: "#FFFFFF",
  surfaceDim: "#E7E3D8",
  line: "#D8D3C4",
  amber: "#D98A1E",
  amberSoft: "#F4E0BC",
  amberText: "#8A5A0E",
  green: "#2D6E5C",
  greenSoft: "#DCEAE5",
  red: "#B23A2E",
  redSoft: "#F4DDD8",
  ink: "#1A2332",
  inkFaint: "#6B7280",
  navy: "#161E2E",
  navySoft: "#212B40",
  onDark: "#F3F1EA",
};

const DARK_COLORS = {
  bg: "#12161F",
  surface: "#1B2130",
  surfaceDim: "#232B3D",
  line: "#333D52",
  amber: "#E0983A",
  amberSoft: "#3A2E15",
  amberText: "#F2B45C",
  green: "#4FAE93",
  greenSoft: "#16302A",
  red: "#E8796A",
  redSoft: "#3A1B18",
  ink: "#EDEBE3",
  inkFaint: "#9AA3B5",
  navy: "#0B0F17",
  navySoft: "#1B2130",
  onDark: "#F3F1EA",
};

const CLIENTS = ["OTIS", "Schindler", "Kone", "TK Elevator", "Mitsubishi", "Fujitec", "Chevalier", "Sigma", "Lecturn", "Hitachi", "Other"];
const DEFAULT_CBM_RATES = {
  Chevalier: 110,
  Fujitec: 165,
  Kone: 162,
  Mitsubishi: 147,
  Schindler: 145,
  Sigma: 160,
  "TK Elevator": 145,
  OTIS: 160,
};
let cbmRateOverrides = {};
function setCbmRateOverridesGlobal(rates) {
  cbmRateOverrides = rates && typeof rates === "object" ? rates : {};
}
function cbmRateFor(client) {
  if (cbmRateOverrides && cbmRateOverrides[client] != null && cbmRateOverrides[client] !== "") {
    const v = Number(cbmRateOverrides[client]);
    if (v >= 0) return v;
  }
  return DEFAULT_CBM_RATES[client] != null ? DEFAULT_CBM_RATES[client] : null;
}

// --- Storage billing: monthly $/CBM rate. First partial month is pro-rated by day;
// every calendar month after that is billed in full regardless of when mid-month
// the goods leave. Billing starts the day after the free-storage period ends. ---
function monthIndexOf(d) { return d.getFullYear() * 12 + d.getMonth(); }
function lastDayOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0); }
function daysInMonthOf(d) { return lastDayOfMonth(d).getDate(); }
function toDateOnly(s) { return new Date(`${s}T00:00:00`); }
function computeStorageCharge(arrivalDateStr, endDateStr, freeDays, ratePerCbmMonth, cbm) {
  if (!arrivalDateStr || !(ratePerCbmMonth > 0) || !(cbm > 0)) return null;
  const arrival = toDateOnly(arrivalDateStr);
  const billStart = new Date(arrival);
  billStart.setDate(billStart.getDate() + (freeDays || 0));
  const end = toDateOnly(endDateStr || todayStr());
  const breakdown = [];
  if (end < billStart) return { billStart: dateToLocalISO(billStart), total: 0, breakdown };
  const firstMonthEnd = lastDayOfMonth(billStart);
  const daysInFirstMonth = daysInMonthOf(billStart);
  let total = 0;
  if (end <= firstMonthEnd) {
    const daysUsed = Math.round((end - billStart) / 86400000) + 1;
    const amt = ratePerCbmMonth * cbm * (daysUsed / daysInFirstMonth);
    breakdown.push({ label: `${fmt(dateToLocalISO(billStart))} \u2013 ${fmt(dateToLocalISO(end))}`, detail: `pro-rata, ${daysUsed}/${daysInFirstMonth} days`, amount: amt, year: billStart.getFullYear(), month: billStart.getMonth() });
    total = amt;
  } else {
    const daysUsed = Math.round((firstMonthEnd - billStart) / 86400000) + 1;
    const amt1 = ratePerCbmMonth * cbm * (daysUsed / daysInFirstMonth);
    breakdown.push({ label: `${fmt(dateToLocalISO(billStart))} \u2013 ${fmt(dateToLocalISO(firstMonthEnd))}`, detail: `pro-rata, ${daysUsed}/${daysInFirstMonth} days`, amount: amt1, year: billStart.getFullYear(), month: billStart.getMonth() });
    total += amt1;
    const numFullMonths = monthIndexOf(end) - monthIndexOf(billStart);
    for (let i = 1; i <= numFullMonths; i++) {
      const mDate = new Date(billStart.getFullYear(), billStart.getMonth() + i, 1);
      const isLast = i === numFullMonths;
      const label = mDate.toLocaleString("en-US", { month: "short", year: "numeric" });
      breakdown.push({ label, detail: isLast && end < lastDayOfMonth(mDate) ? `full month \u2014 left ${fmt(dateToLocalISO(end))}` : "full month", amount: ratePerCbmMonth * cbm, year: mDate.getFullYear(), month: mDate.getMonth() });
      total += ratePerCbmMonth * cbm;
    }
  }
  return { billStart: dateToLocalISO(billStart), total: Math.round(total * 100) / 100, breakdown };
}
// Builds one billing row per arrival batch (or per whole item if no split arrivals were
// used), further split by delivery date where a batch's cases left storage on different
// dates. Rows are not grouped or separated by construction site.
function computeItemBillingRows(item) {
  const rate = cbmRateFor(item.client);
  if (!rate) return [];
  // Oversize is priced per case off that case's own CBM. A lot with #7/29@3.96,
  // #8/29@3.26 and #21/29@5.28 can straddle three tiers, so a single item-wide
  // multiplier would over- or under-charge every case that isn't the one it came from.
  const osMap = oversizeCaseMap(item);
  const namedOversize = osMap.size > 0;
  const lotOversizeCbm = namedOversize ? 0 : oversizeCbmTotal(oversizeCasesOf(item));
  const lotMultiplier = lotOversizeCbm > 0 ? (oversizeTierFor(item.client, lotOversizeCbm) || 1) : 1;
  const tierFor = (cbm) => oversizeTierFor(item.client, cbm) || 1;
  const oversizeMultiplier = namedOversize
    ? Math.max(1, ...[...osMap.values()].map(tierFor))
    : lotMultiplier;
  const effectiveRate = rate * oversizeMultiplier;
  const freeDays = freeDaysFor(item);
  const rows = [];
  if (usesArrivalBatches(item)) {
    const deliveredAt = {};
    activeDeliveries(item).forEach((d) => (d.codes || []).forEach((c) => { deliveredAt[c] = d.date; }));
    const totalPkgCount = (item.packages || []).length || 1;
    const hasCbmData = (item.packages || []).some((p) => Number(p.cbm) > 0);
    activeArrivals(item).forEach((batch) => {
      if (!batch.date) return;
      const byEnd = new Map();
      (batch.codes || []).forEach((code) => {
        const pkg = (item.packages || []).find((p) => p.code === code);
        const osCbm = osMap.get(code);
        let cbm, mult;
        if (osCbm != null) {
          // A named oversize case bills on its declared CBM at its own tier.
          cbm = osCbm;
          mult = tierFor(osCbm);
        } else {
          cbm = pkg ? Number(pkg.cbm) || 0 : 0;
          // No per-case CBM on file (common for older Devan/CFS sheets that only declare an
          // oversize total for the whole lot) - split the declared oversize CBM proportionally
          // across the batch's cases instead of silently billing $0.
          if (!hasCbmData && lotOversizeCbm > 0) cbm = lotOversizeCbm / totalPkgCount;
          mult = namedOversize ? 1 : lotMultiplier;
        }
        const end = deliveredAt[code] || null;
        const key = `${end || "__ongoing__"}|${mult}`;
        if (!byEnd.has(key)) byEnd.set(key, { end, mult, cbm: 0, codes: [] });
        const g = byEnd.get(key);
        g.cbm += cbm;
        g.codes.push(code);
      });
      for (const g of byEnd.values()) {
        if (!(g.cbm > 0)) continue;
        const groupRate = rate * g.mult;
        const calc = computeStorageCharge(batch.date, g.end, freeDays, groupRate, g.cbm);
        if (calc) rows.push({ item, rate: groupRate, baseRate: rate, oversizeMultiplier: g.mult, freeDays, batchDate: batch.date, batchType: batch.type, codes: g.codes, cbm: g.cbm, endDate: g.end, ongoing: !g.end, ...calc });
      }
    });
  } else {
    const arrivalDate = effectiveDepotArrivalDate(item);
    if (!arrivalDate) return [];
    const cbmTotal = (item.packages && item.packages.length) ? item.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0) : Number(item.volumeCbm) || 0;
    if (!(cbmTotal > 0)) return [];
    const status = deriveStatus(item);
    if (status === "delivered") {
      const end = lastDeliveryDate(item);
      const calc = computeStorageCharge(arrivalDate, end, freeDays, effectiveRate, cbmTotal);
      if (calc) rows.push({ item, rate: effectiveRate, baseRate: rate, oversizeMultiplier, freeDays, batchDate: arrivalDate, cbm: cbmTotal, endDate: end, ongoing: false, ...calc });
    } else if (status === "partial") {
      const delivered = new Set(deliveredCodes(item));
      const pkgs = item.packages || [];
      const haveCbmData = pkgs.length > 0 && pkgs.every((p) => p.cbm !== "" && p.cbm != null && !isNaN(Number(p.cbm)));
      let deliveredCbm, remainingCbm, exact;
      if (haveCbmData) {
        deliveredCbm = pkgs.filter((p) => delivered.has(p.code)).reduce((s, p) => s + Number(p.cbm), 0);
        remainingCbm = cbmTotal - deliveredCbm;
        exact = true;
      } else {
        const deliveredShare = deliveredUnits(item) / (totalUnits(item) || 1);
        deliveredCbm = cbmTotal * deliveredShare;
        remainingCbm = cbmTotal - deliveredCbm;
        exact = false;
      }
      const lastDelEnd = lastDeliveryDate(item);
      if (deliveredCbm > 0) {
        const c = computeStorageCharge(arrivalDate, lastDelEnd, freeDays, effectiveRate, deliveredCbm);
        if (c) rows.push({ item, rate: effectiveRate, baseRate: rate, oversizeMultiplier, freeDays, batchDate: arrivalDate, cbm: deliveredCbm, endDate: lastDelEnd, ongoing: false, estimated: !exact, ...c });
      }
      if (remainingCbm > 0) {
        const c = computeStorageCharge(arrivalDate, null, freeDays, effectiveRate, remainingCbm);
        if (c) rows.push({ item, rate: effectiveRate, baseRate: rate, oversizeMultiplier, freeDays, batchDate: arrivalDate, cbm: remainingCbm, endDate: null, ongoing: true, estimated: !exact, ...c });
      }
    } else if (status === "at_depot") {
      const c = computeStorageCharge(arrivalDate, null, freeDays, effectiveRate, cbmTotal);
      if (c) rows.push({ item, rate: effectiveRate, baseRate: rate, oversizeMultiplier, freeDays, batchDate: arrivalDate, cbm: cbmTotal, endDate: null, ongoing: true, ...c });
    }
  }
  return rows;
}
// Aggregates every row's breakdown lines into the given calendar month/year (month is
// 0-indexed, JS Date convention), grouped by client - this is what should match what
// MYOB invoices for that month, for cross-checking.
function computeMonthlyBillingSummary(items, year, month) {
  const byClient = new Map();
  let grandTotal = 0;
  for (const item of items) {
    for (const row of computeItemBillingRows(item)) {
      for (const line of row.breakdown) {
        if (line.year !== year || line.month !== month) continue;
        const amt = Math.round(line.amount * 100) / 100;
        if (!byClient.has(item.client)) byClient.set(item.client, { client: item.client, total: 0, lines: [] });
        const g = byClient.get(item.client);
        g.total += amt;
        g.lines.push({
          project: item.constructionSite || item.project || "",
          jobNumber: item.jobNumber || "",
          batchDate: row.batchDate,
          cbm: row.cbm,
          detail: line.detail,
          label: line.label,
          amount: amt,
          estimated: !!row.estimated,
        });
        grandTotal += amt;
      }
    }
  }
  const clients = [...byClient.values()].map((g) => ({ ...g, total: Math.round(g.total * 100) / 100 })).sort((a, b) => b.total - a.total);
  return { clients, grandTotal: Math.round(grandTotal * 100) / 100 };
}
const ITEM_TYPES = ["Container", "Separate Items"];
const DEPOTS = ["Farspeed Depot 1", "Farspeed Depot 3"];
const DEPOT_LABELS_ZH = {
  "Farspeed Depot 1": "快達一號倉",
  "Farspeed Depot 3": "快達三號倉",
};
function depotLabel(value, lang) {
  if (lang === "zh" && DEPOT_LABELS_ZH[value]) return DEPOT_LABELS_ZH[value];
  return value;
}
function depotDisplay(value, lang) {
  if (!value) return "—";
  if (lang === "zh" && DEPOT_LABELS_ZH[value]) return DEPOT_LABELS_ZH[value];
  return value.replace("Farspeed ", "");
}
// What brings cases into the depot. A Return is goods coming back from site, which is
// an arrival in every way that matters here: it is checked in, it takes up space, and
// storage is charged on it from the day it lands.
const ARRIVING_TYPES = ["Devan", "CFS", "Return"];
const DEFAULT_ROLES = ["Account Officer", "CEO / Business Manager", "Warehouse Depot Head", "Admin"];
const DEFAULT_EMPLOYEES = [
  { id: "E1", name: "Irene Lee", role: "Account Officer" },
  { id: "E2", name: "Nana Chan", role: "Account Officer" },
  { id: "E3", name: "Polly Lee", role: "Account Officer" },
  { id: "E4", name: "Cheng Wai Kee", role: "CEO / Business Manager" },
  { id: "E5", name: "Bhatt Wai Lee", role: "Warehouse Depot Head" },
  { id: "E6", name: "Mega Chan", role: "Admin" },
  { id: "E7", name: "Simon Chan", role: "Staff" },
];
// Login accounts - names match the Employees directory so "Recorded By" and the logged-in
// identity stay consistent. Default password for every account is "Farspeed"; each person
// can change their own password after logging in.
const LOGIN_ACCOUNTS = ["Irene Lee", "Nana Chan", "Cheng Wai Kee", "Mega Chan", "Polly Lee", "Simon Chan", "Bhatt Wai Lee"];
const DEFAULT_PASSWORD = "Farspeed";
async function hashPassword(pw) {
  const enc = new TextEncoder().encode(`farspeed-depot-app:${pw}`);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
const DEFAULT_DIRECTORY = [
  {
    id: "SITE1",
    siteEn: "Kwu Tung North Area 19 Phase 1A & 1B",
    siteZh: "古洞北19區1A & 1B",
    client: "TK Elevator",
    jobRef: "KTN-002",
    orderedBy: "Alex Tam",
    accountOfficer: "Irene Lee",
  },
];
const FREE_DAYS = 14;
const CLIENT_FREE_DAYS = { Schindler: 21 };
let freeStorageRules = [];
function setFreeStorageRulesGlobal(rules) {
  freeStorageRules = Array.isArray(rules) ? rules : [];
}
function freeDaysFor(item) {
  const proj = String(item.project || "").trim().toLowerCase();
  const site = String(item.constructionSite || "").trim().toLowerCase();
  for (const r of freeStorageRules) {
    const key = String(r.project || "").trim().toLowerCase();
    if (!key) continue;
    if ((proj && proj.includes(key)) || (site && site.includes(key))) {
      const d = Number(r.days);
      if (d > 0) return d;
    }
  }
  return CLIENT_FREE_DAYS[item.client] || FREE_DAYS;
}

const FIELD_DEFS = [
  { key: "client", label: "Client", aliases: ["client", "customer"] },
  { key: "project", label: "Project", aliases: ["project", "project name"] },
  { key: "invoiceNumber", label: "Invoice No.", aliases: ["invoice no", "invoice number", "invoice"] },
  { key: "constructionSite", label: "Construction Site Name", aliases: ["construction site", "site", "site name"] },
  { key: "itemType", label: "Item Type", aliases: ["item type", "type"] },
  { key: "packageCount", label: "# of Packages", aliases: ["packages", "no. of packages", "package count", "# of packages"] },
  { key: "unitCode", label: "Escalator/Elevator # or Code", aliases: ["unit code", "escalator/elevator no", "unit no", "code", "unit"] },
  { key: "description", label: "Description", aliases: ["description", "desc"] },
  { key: "shkNumber", label: "Reference / SHK No.", aliases: ["shk no", "shk number", "reference", "reference/shk no"] },
  { key: "weightKg", label: "Weight (KG)", aliases: ["weight", "weight (kg)", "kg"] },
  { key: "volumeCbm", label: "Volume (CBM)", aliases: ["volume", "cbm", "volume (cbm)"] },
  { key: "containers20", label: "No. of 20' Containers", aliases: ["20' containers", "containers 20", "no. of 20' containers", "20ft"] },
  { key: "containers40", label: "No. of 40' Containers", aliases: ["40' containers", "containers 40", "no. of 40' containers", "40ft"] },
  { key: "arrivingType", label: "Arriving Type", aliases: ["arriving type", "arrival type"] },
  { key: "depot", label: "Depot", aliases: ["depot"] },
  { key: "depotLocation", label: "Depot Location / Bay", aliases: ["depot location", "bay", "location"] },
  { key: "terminalArrivalDate", label: "Terminal Arrival Date", aliases: ["terminal arrival date", "terminal arrival"] },
  { key: "terminalLFD", label: "Terminal Last Free Day", aliases: ["terminal lfd", "last free day", "lfd"] },
  { key: "confirmedCollectionDate", label: "Confirmed Collection Date", aliases: ["confirmed collection date", "collection date"] },
  { key: "depotArrivalDate", label: "Depot Arrival Date", aliases: ["depot arrival date", "depot arrival"] },
  { key: "plannedDeliveryDate", label: "Planned Delivery Date", aliases: ["planned delivery date", "delivery date", "estimated delivery"] },
  { key: "jobNumber", label: "Job No.", aliases: ["job no", "job no.", "job number"] },
  { key: "orderedBy", label: "Ordered By", aliases: ["ordered by"] },
  { key: "poNumber", label: "P.O. No.", aliases: ["p.o. no", "po no", "purchase order no"] },
  { key: "jobRef", label: "Job Ref. (Site Code)", aliases: ["job ref", "job ref.", "site code"] },
  { key: "notes", label: "Notes", aliases: ["notes", "remarks"] },
];

function normalizeHeader(h) {
  return String(h || "").toLowerCase().trim().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}
function matchField(header) {
  const norm = normalizeHeader(header);
  const hit = FIELD_DEFS.find((f) => f.aliases.includes(norm) || normalizeHeader(f.label) === norm);
  return hit ? hit.key : null;
}

// --- Client packing-list parser ---------------------------------------
// Different clients (TK Elevator, Schindler, OTIS, etc.) all send packing
// lists with completely different layouts. Rather than expecting a fixed
// header row, this scans for the row that looks most like a table header
// (by keyword match) and maps its columns by meaning, then groups the
// case/package rows underneath by whichever "lot" column is present
// (e.g. a unit code like L0MO-027837.006, or a lift name like L2/L3, or a
// lift no. like "L8 Batch 1") so each lot becomes its own manifest entry.
const PL_HEADER_ALIASES = {
  containerNo: ["container no.", "container no", "container"],
  // Otis heads its case column "(Item)" or "ITEM NO." and numbers them "04#2/3".
  caseNo: ["case no.", "case no", "case\nno", "case", "pkg#", "pkg #", "pkg no.", "pkg no", "package no.", "(item)", "item no.", "item no", "item"],
  // Schindler's own packing lists carry two case columns: a "Case No" running straight
  // through the whole list (1..12 across five lifts) and a "Cases Discript" numbering each
  // lift's own cases ("1/3", "2/3", "3/3"). The job sheets mark the second - "C/S NO. 1-3/3"
  // - so that is the numbering the depot works to, and it wins wherever it carries one.
  caseMark: ["cases discript", "cases     discript", "cases\ndiscript", "case discript", "cases description"],
  qty: ["qty", "quantity", "qyt = quantity", "qyt", "package", "packages", "pkg"],
  lot: ["project no.", "project no", "lift name", "lift no.", "lift no", "lift", "sap no.", "sap no", "sap", "contract no.", "contract no"],
  orderNo: ["omc sales order no.", "omc sales order no", "sales order no.", "sales order no", "order no.", "com.no.", "com no.", "com no", "commission no.", "commission no"],
  description: ["description", "material description"],
  grossWeight: ["g.weight", "gross weight", "gross", "actual   weight", "actual weight", "g.w./kg", "g.w.", "g.w", "gw(kg)", "gw (kg)", "gw"],
  // Schindler's sheets carry an "Estimated weight" beside an "Actual weight". Despite the
  // names, the estimated column is the packing list's gross weight - the sheet says so
  // itself at the top ("the gross weight in the packing list is the theoretic weight") -
  // and it is the column the file's own total row adds up. The actual column is partial:
  // on HPL_0060759755 it reads 517 where estimated reads 800, and totals 29,232 against
  // the sheet's declared 33,736.4. Kept as its own field so it can outrank both.
  estimatedWeight: ["estimated   weight", "estimated  weight", "estimated weight"],
  netWeight: ["n.weight", "net weight", "net", "estimated net weight", "n.w./kg", "n.w.", "n.w", "nw(kg)", "nw (kg)", "nw"],
  // "M3" is how a Chinese factory list heads its volume column - GK230208's WoodenBoxWeight
  // sheet uses it, and without it 305.096 cbm across 125 cases was read as nothing at all.
  // Matching is by whole cell for the short ones: "m3" as a substring would also fire on a
  // dimension column headed "Dimension (M3)" and on stray text elsewhere in a header.
  // Otis heads its volume column "V (M3)", sometimes with fullwidth brackets: "V（m3）".
  cbm: ["cbm", "volume(m3)", "volume (m3)", "volume", "m3", "m\u00b3", "cbm(m3)", "cbm (m3)",
    "\u4f53\u79ef", "\u9ad4\u7a4d", "meas.", "measurement",
    "v(m3)", "v (m3)", "v\uff08m3\uff09", "v\uff08m\u00b3\uff09", "v(m\u00b3)", "v (m\u00b3)"],
  dimension: ["dimension", "dimension (mm)", "dimensions", "size"],
  dimensionCm: ["dim(cm)", "dim (cm)", "dimension(cm)", "dimension (cm)", "dimensions(cm)", "dimensions (cm)"],
};
const PL_CLIENT_HINTS = [
  ["otis", "OTIS"], ["schindler", "Schindler"], ["kone", "Kone"], ["tk elevator", "TK Elevator"],
  ["mitsubishi", "Mitsubishi"], ["fujitec", "Fujitec"], ["chevalier", "Chevalier"], ["sigma", "Sigma"],
  ["lecturn", "Lecturn"], ["hitachi", "Hitachi"],
];

function plNorm(v) {
  return String(v || "").toLowerCase().replace(/\n/g, " ").replace(/\s+/g, " ").trim();
}
function plNum(v) {
  if (v === "" || v === null || v === undefined) return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const cleaned = String(v).replace(/,/g, "").trim();
  const n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}
// The row at the foot of a packing list that adds the sheet up is not a case, and taking it
// for one puts the whole sheet's weight and volume onto whichever lot happened to be last.
// Only the English word was looked for before, so a mainland factory sheet ending in
// "合  计" - Chinese, and spaced out to fill the cell - went straight through as a case.
// A subtotal in the middle of a sheet is skipped rather than treated as the end of the table,
// since real cases follow it.
const PL_TOTAL_LABELS = ["\u5408\u8a08", "\u5408\u8a08", "\u5408\u8ba1", "\u603b\u8ba1", "\u7e3d\u8a08", "\u5171\u8ba1", "\u5171\u8a08", "\u603b\u6570", "\u7e3d\u6578", "\u5408\u5171", "\u7e3d\u91cd", "\u603b\u91cd"];
const PL_SUBTOTAL_LABELS = ["\u5c0f\u8ba1", "\u5c0f\u8a08"];
function plTotalsRowKind(row) {
  let kind = "";
  for (const cell of row || []) {
    const text = plNorm(cell);
    if (!text) continue;
    // "合  计" is written with padding inside the cell, so spacing is dropped before matching.
    const compact = text.replace(/\s+/g, "");
    if (PL_SUBTOTAL_LABELS.some((l) => compact.startsWith(l))) return "subtotal";
    if (PL_TOTAL_LABELS.some((l) => compact.startsWith(l))) kind = "total";
    // The original reading, kept exactly as it was: any cell mentioning "total" anywhere
    // ends the table, sub-totals included.
    if (text.includes("total")) kind = "total";
  }
  return kind;
}
// Mitsubishi's factory list announces which lift a case belongs to on the case's own row,
// in parentheses beside the description - "(#.01)", "(#.09)". That number is the first part
// of the marking painted on the box, and the list is the only document in the chain that
// leaves it out of the case column.
function plLiftMarkOnRow(row) {
  for (const cell of row || []) {
    const m = String(cell == null ? "" : cell).match(/\(\s*#\s*\.?\s*(\d{1,3})\s*\)/);
    if (m) return m[1];
  }
  return "";
}
// The component half of a split Mitsubishi marking is always one letter and two digits -
// B11, E21, D51, A10, Z11 - optionally with a letter after it. Nothing is rebuilt unless
// the case number looks like that. Without the check, a "(#.3)" written in passing in some
// other maker's description column would rebuild that row's ordinary "1/17" into "31/17".
const PL_MITSUBISHI_CASE_RE = /^[A-Za-z]\d{2}[A-Za-z]?(?=$|[\s\-0-9])/;
function plMitsubishiMarking(caseNo, liftMark) {
  const compact = String(caseNo || "").replace(/\s+/g, "");
  if (!liftMark || !compact) return caseNo;
  if (!PL_MITSUBISHI_CASE_RE.test(compact)) return caseNo;
  if (compact.startsWith(liftMark)) return caseNo;
  return `${liftMark}${compact}`;
}
function plCbmFromDimension(v, unit) {
  // Parses strings like "2000*850*600" (mm) or "140*90*96" (cm) into m3.
  const parts = String(v || "").split(/[x*×]/i).map((s) => Number(s.replace(/,/g, "").trim()));
  if (parts.length !== 3 || parts.some((n) => !n || isNaN(n))) return 0;
  const divisor = unit === "cm" ? 1e6 : 1e9;
  return (parts[0] * parts[1] * parts[2]) / divisor;
}
function plScoreRow(row) {
  let score = 0;
  row.forEach((cell) => {
    const n = plNorm(cell);
    if (!n) return;
    for (const aliases of Object.values(PL_HEADER_ALIASES)) {
      if (aliases.some((a) => n === a || n.includes(a))) { score++; break; }
    }
  });
  return score;
}
function plDetectHeaderRow(rows) {
  let bestIdx = -1, bestScore = 0;
  rows.slice(0, 40).forEach((row, idx) => {
    const s = plScoreRow(row);
    if (s > bestScore) { bestScore = s; bestIdx = idx; }
  });
  return bestScore >= 3 ? bestIdx : -1;
}
function plMapColumns(headerRow) {
  const colMap = {};
  headerRow.forEach((cell, idx) => {
    const n = plNorm(cell);
    if (!n) return;
    for (const [field, aliases] of Object.entries(PL_HEADER_ALIASES)) {
      if (colMap[field] !== undefined) continue;
      // A short alias has to match the whole cell. "m3" or "net" appearing inside a longer
      // heading is far more often part of another column's name than that column itself.
      const hit = aliases.some((a) => (a.length <= 3 ? n === a : n === a || n.includes(a)));
      if (hit) { colMap[field] = idx; break; }
    }
  });
  return colMap;
}
// Manufacturer packing lists (like TK Elevator's) often list a "Marks:" block where the
// first line is the project/building name (sometimes with unit codes and city tacked on,
// e.g. "1881 Heritage E3&4,HongKong") followed by lines pairing a part/SAP number with its
// unit code (e.g. "410217-501-003/E3"). This pulls out both: the clean project name, and a
// lookup from part number to unit code so case groups can be labeled "E3"/"E4" instead of
// showing the raw part number.
function plGuessMarksBlock(rows) {
  const legend = {};
  let rawProjectLine = "";
  for (let r = 0; r < Math.min(rows.length, 25); r++) {
    const row = rows[r];
    for (let i = 0; i < row.length; i++) {
      const n = plNorm(row[i]);
      if (n !== "marks" && n !== "marks:" && n !== "唛頭" && n !== "唛头" && n !== "shipping marks") continue;
      let valueCol = -1;
      for (let j = i + 1; j < row.length; j++) {
        if (row[j] !== "" && row[j] != null) { rawProjectLine = String(row[j]).trim(); valueCol = j; break; }
      }
      if (valueCol === -1 && rows[r + 1] && row[i] !== undefined) {
        if (rows[r + 1][i] !== "" && rows[r + 1][i] != null) { rawProjectLine = String(rows[r + 1][i]).trim(); valueCol = i; r = r + 1; }
      }
      if (valueCol === -1) continue;
      for (let rr = r + 1; rr < Math.min(r + 6, rows.length); rr++) {
        const cell = rows[rr] ? rows[rr][valueCol] : undefined;
        if (cell === "" || cell == null) break;
        const text = String(cell).trim();
        const m = text.match(/^([\w\-]+)\s*\/\s*([A-Za-z0-9]+)$/);
        if (!m) break;
        legend[m[1].trim()] = m[2].trim();
      }
      return { rawProjectLine, legend };
    }
  }
  return { rawProjectLine, legend };
}
function plCleanProjectName(rawLine) {
  if (!rawLine) return "";
  // Drop a trailing unit-code list and city (e.g. "1881 Heritage E3&4,HongKong" -> "1881 Heritage")
  const beforeComma = rawLine.split(",")[0];
  const stripped = beforeComma.replace(/\s*[A-Za-z]\d+(\s*&\s*[A-Za-z]?\d+)*\s*$/i, "").trim();
  return stripped || beforeComma.trim();
}
function plGuessClient(rows) {
  for (const row of rows.slice(0, 25)) {
    for (const cell of row) {
      const n = plNorm(cell);
      if (!n) continue;
      for (const [needle, label] of PL_CLIENT_HINTS) {
        if (n.includes(needle)) return label;
      }
    }
  }
  return null;
}
function plGuessProject(rows) {
  // "Project Name" can appear two ways. On a job-sheet-style layout it's a label with the
  // value in the cell beside it. On Schindler's booking workbook it's a column header in
  // the table, with "Destination Port" as the next header along - reading rightwards
  // there returns the neighbouring header instead of a project. So when the label sits in
  // a detected header row, take the first real value from underneath it.
  const headerIdx = plDetectHeaderRow(rows);
  if (headerIdx !== -1) {
    const headerRow = rows[headerIdx] || [];
    for (let i = 0; i < headerRow.length; i++) {
      const n = plNorm(headerRow[i]);
      if (n !== "project" && n !== "project name" && n !== "project name:") continue;
      for (let r = headerIdx + 1; r < rows.length; r++) {
        const v = String((rows[r] || [])[i] || "").trim();
        if (v) return v;
      }
    }
  }
  for (const row of rows.slice(0, 25)) {
    for (let i = 0; i < row.length; i++) {
      const n = plNorm(row[i]);
      if (n === "project" || n === "project:" || n === "project name" || n === "project name:") {
        for (let j = i + 1; j < row.length; j++) {
          if (row[j] !== "" && row[j] !== null && row[j] !== undefined) return String(row[j]).trim();
        }
      }
    }
  }
  // Fallback: manufacturer packing lists often have no literal "Project" label at all -
  // pull the cleaned building name out of the Marks block instead.
  const { rawProjectLine } = plGuessMarksBlock(rows);
  return plCleanProjectName(rawProjectLine);
}
function parsePackingListSheet(rows, legend) {
  const headerIdx = plDetectHeaderRow(rows);
  if (headerIdx === -1) return null;
  const colMap = plMapColumns(rows[headerIdx]);
  if (colMap.lot === undefined && colMap.caseNo === undefined && colMap.caseMark === undefined) return null;

  // Some workbooks give a sheet per lot rather than a lot column: Otis names the contract
  // and lift in a labelled cell above the table - "Contract No. | 51N02002/L1" - and every
  // case on the sheet belongs to it. Read once, and used only where no lot column exists.
  let sheetLot = "";
  if (colMap.lot === undefined) {
    for (let r = 0; r < Math.min(headerIdx, rows.length); r++) {
      const cells = rows[r] || [];
      for (let c = 0; c < cells.length; c++) {
        const n = plNorm(cells[c]).replace(/[:\uff1a]/g, "").trim();
        if (!["contract no.", "contract no", "\u5408\u540c\u53f7"].includes(n)) continue;
        for (let cc = c + 1; cc < Math.min(c + 5, cells.length); cc++) {
          const v = String(cells[cc] == null ? "" : cells[cc]).trim();
          if (v) { sheetLot = v; break; }
        }
        if (sheetLot) break;
      }
      if (sheetLot) break;
    }
  }

  // Translates a raw lot value (which may be a manufacturer part/SAP number, or several
  // comma-separated ones for cases shared across units) into friendly unit codes via the
  // legend pulled from the Marks block, e.g. "410217-501-003,502-004" -> "E3, E4". The
  // second part often omits the shared numeric prefix, so that's restored first.
  function expandAbbreviatedParts(raw) {
    const parts = raw.split(",").map((p) => p.trim()).filter(Boolean);
    if (parts.length <= 1) return parts;
    const firstSegs = parts[0].split("-");
    const prefix = firstSegs.length > 2 ? firstSegs.slice(0, firstSegs.length - 2).join("-") : null;
    return parts.map((p, idx) => {
      if (idx === 0) return p;
      const segs = p.split("-");
      if (prefix && segs.length < firstSegs.length) return `${prefix}-${p}`;
      return p;
    });
  }
  function translateLot(raw) {
    if (!legend || !Object.keys(legend).length) return raw;
    const parts = expandAbbreviatedParts(raw);
    const translated = parts.map((p) => legend[p] || p);
    const allTranslated = parts.every((p) => legend[p]);
    return allTranslated ? [...new Set(translated)].join(", ") : raw;
  }

  const groups = {};
  const order = [];
  let lastLot = "", lastContainer = "", lastCase = "", lastOrderNo = "";
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (row.every((c) => String(c || "").trim() === "")) continue;
    const totalsKind = plTotalsRowKind(row);
    if (totalsKind === "subtotal") continue;
    if (totalsKind === "total") break;

    const lot = colMap.lot !== undefined ? String(row[colMap.lot] || "").trim() : "";
    const container = colMap.containerNo !== undefined ? String(row[colMap.containerNo] || "").trim() : "";
    // The per-lift marking wins where the row carries one; the running case number is the
    // fallback for lists that only have the one column.
    const caseMark = colMap.caseMark !== undefined ? String(row[colMap.caseMark] || "").trim() : "";
    const plainCase = colMap.caseNo !== undefined ? String(row[colMap.caseNo] || "").trim() : "";
    const readCase = /^\s*\d+\s*\/\s*\d+\s*$/.test(caseMark) ? caseMark.replace(/\s+/g, "") : (plainCase || caseMark);
    // Where the row names its lift and the case number is a split Mitsubishi marking, the
    // two are put back together: the list's own columns leave "B11 01" and "E21 01A" where
    // the Delivery Memo, the CFS sheet and the delivery job sheet all write 01B1101 and
    // 01E2101A. A delivery asking for the marking on the box found nothing at the depot
    // and pre-selected nothing. Any other maker's case number is left exactly as it is.
    const caseNo = plMitsubishiMarking(readCase, plLiftMarkOnRow(row));
    const orderNo = colMap.orderNo !== undefined ? String(row[colMap.orderNo] || "").trim() : "";
    const description = colMap.description !== undefined ? String(row[colMap.description] || "").trim() : "";
    const grossVal = colMap.grossWeight !== undefined && row[colMap.grossWeight] !== "" && row[colMap.grossWeight] != null ? plNum(row[colMap.grossWeight]) : null;
    const netVal = colMap.netWeight !== undefined && row[colMap.netWeight] !== "" && row[colMap.netWeight] != null ? plNum(row[colMap.netWeight]) : null;
    const estVal = colMap.estimatedWeight !== undefined && row[colMap.estimatedWeight] !== "" && row[colMap.estimatedWeight] != null ? plNum(row[colMap.estimatedWeight]) : null;
    // Schindler's sheets carry an Estimated weight beside an Actual weight, and the depot
    // works to whichever of the two is heavier. That choice is made per lot on the column
    // totals, not row by row: taking the larger figure case by case mixes the two measures
    // and overshoots both - on this lot it would read 8,702.04 against an actual column of
    // 8,404.84 and an estimated column of 7,517.48. The columns are summed below and the
    // heavier one wins whole, so 60717017 reads 8,404.84 while L32-02, whose estimated
    // column is the heavier at 9,254.6, keeps that.
    //
    // Where there is no estimated column, use whichever of gross/net is bigger, which
    // doesn't assume gross is.
    const weight = estVal != null && estVal > 0
      ? estVal
      : (grossVal != null && netVal != null ? Math.max(grossVal, netVal) : (grossVal != null ? grossVal : (netVal != null ? netVal : 0)));
    let cbm = 0;
    if (colMap.dimensionCm !== undefined && row[colMap.dimensionCm]) cbm = plCbmFromDimension(row[colMap.dimensionCm], "cm");
    if (!cbm && colMap.dimension !== undefined && row[colMap.dimension]) cbm = plCbmFromDimension(row[colMap.dimension]);
    if (!cbm && colMap.cbm !== undefined && row[colMap.cbm] !== "" && row[colMap.cbm] != null) cbm = plNum(row[colMap.cbm]);

    // Many bilingual packing lists have a second header row directly below the first,
    // repeating the same column labels in Chinese (e.g. "项目号/行号/梯号" under "SAP NO.").
    // That row has no real weight or CBM data, so skip it rather than counting it as a case.
    if (i === headerIdx + 1 && !weight && !cbm) continue;

    // A blank lift cell means "same lift as the row above" only while we're still inside
    // the same order. On a mixed booking sheet the lift column is filled for the lift's
    // own cases and left blank for the spare-parts orders around them - carrying the lift
    // forward there swept every later spare part into it, turning L34-S1's nine cases
    // into nineteen. A new order number ends the run.
    if (orderNo && lastOrderNo && orderNo !== lastOrderNo) lastLot = "";
    if (lot) lastLot = lot;
    if (container) lastContainer = container;
    if (caseNo) lastCase = caseNo;
    if (orderNo) lastOrderNo = orderNo;
    // A description usually signals a real data row, but some packing lists (like ones
    // that only list PKG#/dimensions/weight) have no description column at all - a case
    // number paired with real weight or CBM is equally good evidence of a genuine row.
    if (!description && !(caseNo && (weight || cbm))) continue;

    // Nested layouts list a case once and then itemise what's inside it on the rows
    // below - Schindler's "Detail" sheet gives case 1/17 its volume and weight, then
    // three bolt lines underneath with no case number, no volume and no weight of their
    // own. Those lines are contents, not cases: counting them turned 17 cases into 68.
    // A row that names no case and carries neither weight nor volume belongs to the case
    // above it, and would contribute nothing but an inflated package count anyway.
    if (!caseNo && !weight && !cbm && lastCase) continue;

    // Schindler's booking workbooks identify a lot by its OMC Sales Order no. Where a lift
    // column exists it wins, but rows without a lift still belong to their own order -
    // lumping them together under UNSPECIFIED merged unrelated spare-parts consignments
    // into one entry.
    // Where a list names no lift and no order, the case number often carries the lift
    // itself - GK230208 numbers its cases "L_P01/G1", "L_W05/G23", ten lifts across 125
    // cases - and without reading that the whole file lands as one entry called
    // UNSPECIFIED. The prefix must contain a letter to be taken as a lift: splitting on
    // the part before the slash in "1/24" or "3/19" would shred an ordinary case list
    // into a lot per case.
    const casePrefix = /[A-Za-z]/.test(String(lastCase || "").split("/")[0] || "")
      ? String(lastCase).split("/")[0].trim()
      : "";
    const key = translateLot(lastLot) || sheetLot || lastOrderNo || casePrefix || "UNSPECIFIED";
    if (!groups[key]) { groups[key] = { lot: key, packages: [], containers: new Set(), totalWeight: 0, totalCbm: 0 }; order.push(key); }
    groups[key].packages.push({
      code: lastCase || String(groups[key].packages.length + 1), orderNo: lastOrderNo, description,
      weightKg: weight ? String(weight) : "", cbm: cbm ? String(cbm) : "",
      // Both readings are kept on the case so the heavier column can be picked per lot
      // once every row has been read. Stripped again below.
      __est: estVal != null && estVal > 0 ? estVal : null,
      __act: grossVal != null && grossVal > 0 ? grossVal : null,
    });
    if (lastContainer) groups[key].containers.add(lastContainer);
    groups[key].totalWeight += weight;
    groups[key].totalCbm += cbm;
  }
  // A single lot can combine multiple separate order numbers, each restarting its own case
  // numbering (e.g. two different orders both having a case "1/1") - disambiguate any case
  // code that collides within the same lot so every package has a genuinely unique code.
  for (const key of order) {
    const pkgs = groups[key].packages;
    const counts = {};
    pkgs.forEach((p) => { counts[p.code] = (counts[p.code] || 0) + 1; });
    const seen = {};
    for (const p of pkgs) {
      if (counts[p.code] > 1 && p.orderNo) {
        seen[p.code] = (seen[p.code] || 0) + 1;
        p.code = `${p.orderNo}/${p.code}`;
      }
    }
  }
  // Now that every row is in, settle the Estimated-vs-Actual choice per lot on the column
  // totals. A lot whose rows carry both readings takes the heavier column whole; where a
  // case is missing a figure in the winning column, its other reading stands in rather
  // than leaving the case at zero.
  for (const key of order) {
    const pkgs = groups[key].packages;
    let estTotal = 0, actTotal = 0, both = false;
    for (const p of pkgs) {
      if (p.__est != null) estTotal += p.__est;
      if (p.__act != null) actTotal += p.__act;
      if (p.__est != null && p.__act != null) both = true;
    }
    if (both && estTotal > 0 && actTotal > 0) {
      const useActual = actTotal > estTotal;
      for (const p of pkgs) {
        const chosen = useActual ? (p.__act != null ? p.__act : p.__est) : (p.__est != null ? p.__est : p.__act);
        p.weightKg = chosen ? String(chosen) : "";
      }
      groups[key].totalWeight = pkgs.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
    }
    for (const p of pkgs) { delete p.__est; delete p.__act; }
  }
  // A shipper's lift column is not always right. On the 60766021/60766022 booking, four
  // rows carry the wrong lift: cases 4/20 and 6/20 are labelled L52 though their order is
  // 60766022 (L53's), and cases 18/19 and 19/19 are labelled L53 though their order is
  // 60766021 (L52's). Grouping on the lift alone put four boxes in the wrong lot, so the
  // depot showed L52 holding a 4/20 it never received and L53 short of the cases the
  // delivery sheet asked for.
  //
  // The order number is the reliable identifier - every case of order 60766021 is numbered
  // out of 19 and every case of 60766022 out of 20, which the case codes themselves
  // confirm - so where an order sits overwhelmingly under one lift, the stragglers filed
  // under another are moved back to it. A lift that legitimately combines several orders
  // is untouched: each of its orders appears under that lift only, so none has a majority
  // anywhere else.
  if (colMap.orderNo !== undefined && colMap.lot !== undefined) {
    const homeOf = new Map(); // orderNo -> lot key holding most of its cases
    const counts = new Map(); // orderNo -> Map(lot key -> count)
    for (const key of order) {
      for (const p of groups[key].packages) {
        if (!p.orderNo) continue;
        if (!counts.has(p.orderNo)) counts.set(p.orderNo, new Map());
        const m = counts.get(p.orderNo);
        m.set(key, (m.get(key) || 0) + 1);
      }
    }
    for (const [orderNo, m] of counts) {
      if (m.size < 2) continue;
      const ranked = [...m.entries()].sort((a, b) => b[1] - a[1]);
      // Only a clear majority moves anything - a genuine even split is left alone for a
      // person to look at rather than guessed at.
      if (ranked[0][1] > ranked[1][1]) homeOf.set(orderNo, ranked[0][0]);
    }
    const touched = new Set();
    for (const key of order) {
      const stay = [];
      for (const p of groups[key].packages) {
        const home = p.orderNo ? homeOf.get(p.orderNo) : null;
        if (!home || home === key || !groups[home]) { stay.push(p); continue; }
        groups[home].packages.push(p);
        touched.add(home);
        touched.add(key);
      }
      if (stay.length !== groups[key].packages.length) groups[key].packages = stay;
    }
    for (const key of touched) {
      const g = groups[key];
      // Sort on the lot size first so that if a group does end up holding two case series
      // they stay in their own runs rather than interleaving 1/2 between 1/3 and 2/3.
      const denom = (c) => { const m = String(c).match(/\/(\d+)\s*$/); return m ? Number(m[1]) : 0; };
      g.packages.sort((a, b) => denom(a.code) - denom(b.code) || codeLeadingNumber(a.code) - codeLeadingNumber(b.code));
      g.totalWeight = g.packages.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
      g.totalCbm = g.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0);
    }
  }
  // Where the order number stood in as the lot, repeating it as a per-package order
  // heading would just print "60789730 / 60789730" above its own cases.
  for (const key of order) {
    for (const p of groups[key].packages) if (p.orderNo && p.orderNo === key) p.orderNo = "";
  }
  const identifiedLots = colMap.lot !== undefined
    || (colMap.orderNo !== undefined && order.length > 0 && !order.includes("UNSPECIFIED"));
  return { groups: order.map((k) => ({ ...groups[k], containers: [...groups[k].containers] })), hasLotColumn: identifiedLots };
}
// Last resort for a workbook whose case-level pages can't be read: the Summary page.
// It states the totals directly - "Total No. Case : 17", "Gross Weight : 16735",
// "Volume : 30.29" - so a lot can still be created with the right case count and totals
// even when no per-case table was found. Per-case weight and volume are shared evenly,
// which is an estimate, but keeps storage billing (which is charged on per-case CBM)
// working instead of silently pricing the lot at zero.
// Mitsubishi's factory packing list (装箱清单) is laid out nothing like Schindler's: no
// column headers over a case table, but a header block stating the totals in Chinese -
// 总净重 net weight, 总毛重 gross weight, 总体积 volume, 总箱数 case count - above a long
// listing where each case is announced by a two-part number in the 箱号/CASE NUMBER column
// ("A21" then "01", "C31" then "01-4-2") and everything under it is that case's contents.
//
// So the cases are counted from those announcements and checked against the stated 总箱数,
// and the weight and volume come off the header rather than being summed. The gross weight
// is the one taken: that is what the depot stores and bills on.
function plCjkLabelledNumber(rows, needles) {
  for (const row of rows || []) {
    for (let i = 0; i < row.length; i++) {
      const cell = String(row[i] == null ? "" : row[i]);
      if (!needles.some((n) => cell.includes(n))) continue;
      for (let j = i + 1; j < row.length; j++) {
        const v = plNum(row[j]);
        if (v > 0) return v;
      }
    }
  }
  return 0;
}
function plCjkLabelledText(rows, needles) {
  for (const row of rows || []) {
    for (let i = 0; i < row.length; i++) {
      const cell = String(row[i] == null ? "" : row[i]);
      if (!needles.some((n) => cell.includes(n))) continue;
      for (let j = i + 1; j < row.length; j++) {
        const v = String(row[j] == null ? "" : row[j]).trim();
        if (v) return v;
      }
    }
  }
  return "";
}
function parseMitsubishiPackingList(rows) {
  const cases = Number(plCjkLabelledNumber(rows, ["\u603b\u7bb1\u6570", "packing amount"]));
  const gross = plCjkLabelledNumber(rows, ["\u603b\u6bdb\u91cd", "gross weight"]);
  const net = plCjkLabelledNumber(rows, ["\u603b\u51c0\u91cd", "net weight"]);
  const cbm = plCjkLabelledNumber(rows, ["\u603b\u4f53\u79ef", "cubicmeter", "volume,cubic"]);
  if (!(cases > 0) || !(gross > 0 || net > 0)) return null;

  // Find the 箱号/CASE NUMBER column, then read the case identifiers announced under it.
  let headerRow = -1, caseCol = -1;
  for (let r = 0; r < rows.length && headerRow < 0; r++) {
    for (let c = 0; c < (rows[r] || []).length; c++) {
      const cell = String(rows[r][c] == null ? "" : rows[r][c]);
      if (cell.includes("\u7bb1\u53f7") || /case\s*number/i.test(cell)) { headerRow = r; caseCol = c; break; }
    }
  }
  const codes = [];
  // The lift a case belongs to is announced on the same row in parentheses - "(#.09)" -
  // and it is the first part of the marking painted on the box.
  const liftOnRow = (row) => {
    for (const cell of row || []) {
      const m = String(cell == null ? "" : cell).match(/\(\s*#\s*\.?\s*(\d{1,3})\s*\)/);
      if (m) return m[1];
    }
    return "";
  };
  if (caseCol >= 0) {
    const skip = ["\u6346\u5305\u5305\u88c5", "BUNDLE", "CASE", "\uff08\u5c01\u95ed\uff09\u6728\u7bb1", "\u7bb1\u53f7"];
    for (let r = headerRow + 1; r < rows.length; r++) {
      const raw = String((rows[r] || [])[caseCol] == null ? "" : rows[r][caseCol]).trim();
      if (!raw || skip.some((k) => raw.includes(k)) || /case\s*number/i.test(raw)) continue;
      // The case number is split across two columns: a type prefix and a running number.
      let suffix = "";
      for (let c = caseCol + 1; c < Math.min(caseCol + 6, (rows[r] || []).length); c++) {
        const v = String(rows[r][c] == null ? "" : rows[r][c]).trim();
        if (v) { suffix = v; break; }
      }
      if (!suffix) continue;
      // The factory list splits the marking into columns; every other document in the
      // chain - the Delivery Memo, the CFS sheet, the delivery job sheet - writes it whole
      // as lift, type, running number: "09" + "B11" + "09" is 09B1109, and "01" + "E21" +
      // "01A" is 01E2101A. Joining the two columns with a hyphen instead produced "B11-09",
      // a marking that appears on no piece of paper and on no box, so a delivery asking for
      // 01B1101 found nothing at the depot and pre-selected nothing.
      const lift = liftOnRow(rows[r]);
      const code = lift && PL_MITSUBISHI_CASE_RE.test(raw) ? `${lift}${raw}${suffix}` : `${raw}-${suffix}`;
      if (!codes.includes(code)) codes.push(code);
    }
  }
  // The stated count is what the factory certifies; where the listing disagrees, take the
  // stated one and number the remainder plainly rather than quietly shipping a short list.
  const list = codes.length === cases ? codes : Array.from({ length: cases }, (_, i) => `${i + 1}/${cases}`);
  const weight = gross > 0 ? gross : net;
  const per = (total, i) => {
    if (!total) return "";
    const each = Math.round((total / list.length) * 1000) / 1000;
    return String(i === list.length - 1 ? Math.round((total - each * (list.length - 1)) * 1000) / 1000 : each);
  };
  const description = plCjkLabelledText(rows, ["\u4ea7\u54c1\u540d\u79f0", "product name"]) || "ELEVATOR PARTS";
  // The shipping mark - the leading number of 合同号, "1325003000 ZS1680-260350-OV0" - is
  // the one identifier this factory list shares with the Delivery Memo that follows it, so
  // it is the lot. The DM's own number is what these lots are ultimately filed under, but
  // it does not exist yet when the factory list is written.
  const contract = plCjkLabelledText(rows, ["\u5408\u540c\u53f7", "contract number"]);
  const mark = (String(contract).trim().split(/\s+/)[0] || "").trim();
  const lot = mark || plCjkLabelledText(rows, ["\u8ba2\u5355\u53f7", "order number"]) || "";
  return {
    groups: [{
      lot: lot || "UNSPECIFIED",
      containers: [],
      totalWeight: weight,
      totalCbm: cbm,
      packages: list.map((code, i) => ({
        code, orderNo: "", description,
        weightKg: per(weight, i), cbm: per(cbm, i),
      })),
    }],
    client: plCjkLabelledText(rows, ["\u987e\u5ba2\u540d\u79f0", "consignee"]),
    project: plCjkLabelledText(rows, ["\u5927\u697c\u540d\u79f0", "building name"]),
  };
}
function parsePackingListSummarySheet(rows) {
  const labelled = (labels) => {
    for (const row of rows || []) {
      for (let i = 0; i < row.length; i++) {
        const n = plNorm(row[i]).replace(/[:\uff1a]/g, "").trim();
        if (!labels.includes(n)) continue;
        for (let j = i + 1; j < row.length; j++) {
          const v = plNum(row[j]);
          if (v > 0) return v;
        }
      }
    }
    return 0;
  };
  const cases = Math.round(labelled(["total no. case", "total no of case", "total no. of case", "total cases", "total case"]));
  if (!(cases > 0) || cases > 5000) return null;
  const gross = labelled(["gross weight", "total gross weight", "g.w."]);
  const net = labelled(["total net weight", "net weight", "n.w."]);
  const volume = labelled(["volume", "total volume", "measurement"]);
  const weight = gross > 0 ? gross : net;
  if (!(weight > 0) && !(volume > 0)) return null;

  let lot = "";
  for (const row of rows || []) {
    for (let i = 0; i < row.length; i++) {
      const n = plNorm(row[i]).replace(/[:\uff1a]/g, "").trim();
      if (n !== "project name" && n !== "project") continue;
      for (let j = i + 1; j < row.length; j++) {
        const v = String(row[j] == null ? "" : row[j]).trim();
        if (v) { lot = v; break; }
      }
      if (lot) break;
    }
    if (lot) break;
  }
  // "PL_4550362030_AST(L3401-L3406)" - the lifts are the part in brackets.
  const bracketed = lot.match(/\(([^)]+)\)\s*$/);
  if (bracketed) lot = bracketed[1].trim();

  const packages = Array.from({ length: cases }, (_, i) => ({
    code: `${i + 1}/${cases}`,
    orderNo: "",
    description: "",
    weightKg: weight > 0 ? String(Math.round((weight / cases) * 100) / 100) : "",
    cbm: volume > 0 ? String(Math.round((volume / cases) * 10000) / 10000) : "",
  }));
  return {
    groups: [{
      lot: lot || "UNSPECIFIED",
      packages,
      containers: [],
      // Totals come off the Summary itself, so they stay exact rather than inheriting
      // the rounding drift of the per-case shares above.
      totalWeight: weight,
      totalCbm: volume,
      fromSummary: true,
    }],
    hasLotColumn: !!lot,
  };
}
function parsePackingListWorkbook(workbook) {
  // A workbook may be organised one sheet per lot, with no single sheet holding the whole
  // shipment: Otis's CLD-51N02002 has a summary page, two container pages, and one page per
  // contract naming its own lift - "Contract No. 51N02002/L1". Picking a single best sheet
  // returns one contract of four, and the container pages carry the same cases again under
  // the same contract numbers, so merging everything would double them.
  //
  // So a per-lot workbook is recognised first: two or more sheets that each name their own
  // contract and hold their own cases. Those are merged, and nothing else is read.
  const perLot = [];
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
    const { legend } = plGuessMarksBlock(rows);
    const result = parsePackingListSheet(rows, legend);
    if (!result || !result.groups || result.groups.length !== 1) continue;
    if (result.hasLotColumn) continue;
    const lot = result.groups[0].lot;
    if (!lot || lot === "UNSPECIFIED") continue;
    perLot.push({ sheetName, rows, group: result.groups[0] });
  }
  const perLotNames = perLot.map((x) => x.group.lot);
  if (perLot.length >= 2 && new Set(perLotNames).size === perLot.length) {
    let c = null, pj = "";
    for (const x of perLot) {
      if (!c) c = plGuessClient(x.rows);
      if (!pj) pj = plGuessProject(x.rows);
    }
    return { groups: perLot.map((x) => x.group), client: c, project: pj };
  }
  let bestGroups = null;
  let bestHasLotColumn = false;
  let client = null;
  let project = "";
  // Schindler workbooks pair a case-level "packing list_Head" sheet with a much longer
  // "Packing List Contents" sheet that breaks every case down into its component lines.
  // Both parse, but only the Head sheet is one row per case - Contents would turn twenty
  // cases into four hundred packages - so it wins outright wherever it exists.
  let headSheetUsed = false;
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    if (!client) client = plGuessClient(rows);
    if (!project) project = plGuessProject(rows);
    // The Mitsubishi 装箱清单 is tried first: its header block states 总箱数 and 总毛重
    // outright, which is a far stronger signal than the generic reader's column sniffing -
    // left to itself that reader mistakes the 箱号 column for a case table and returns one
    // "package" per announcement line, weightless, at roughly twice the true case count.
    const mitsubishi = parseMitsubishiPackingList(rows);
    if (mitsubishi) {
      bestGroups = mitsubishi.groups;
      bestHasLotColumn = true;
      if (!client) client = mitsubishi.client;
      if (!project) project = mitsubishi.project;
      break;
    }
    const { legend } = plGuessMarksBlock(rows);
    const result = parsePackingListSheet(rows, legend);
    if (!result || !result.groups || result.groups.length === 0) continue;
    const { groups, hasLotColumn } = result;
    const isHeadSheet = /packing\s*list[\s_\-]*head/i.test(sheetName);
    if (isHeadSheet && !headSheetUsed) {
      bestGroups = groups; bestHasLotColumn = hasLotColumn; headSheetUsed = true;
      continue;
    }
    if (headSheetUsed) continue;
    // A sheet that actually identifies lift/lot numbers is preferred over one that had to
    // lump everything into a single UNSPECIFIED group, even if the latter has more raw
    // rows (that's often a material/component breakdown sheet, not the case-level one).
    const better = !bestGroups || (hasLotColumn && !bestHasLotColumn) || (hasLotColumn === bestHasLotColumn && groups.length > bestGroups.length);
    if (better) { bestGroups = groups; bestHasLotColumn = hasLotColumn; }
  }
  // Nothing case-level anywhere in the workbook - fall back to the Summary page.
  if (!bestGroups || bestGroups.length === 0) {
    for (const sheetName of workbook.SheetNames) {
      if (!/summary|\u603b\u8ba1|\u6458\u8981/i.test(sheetName)) continue;
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: true });
      const summary = parsePackingListSummarySheet(rows);
      if (summary) { bestGroups = summary.groups; break; }
    }
  }
  return { groups: bestGroups, client, project };
}

function sigPart(v) {
  return String(v || "").trim().toLowerCase();
}
function itemSignature(item) {
  const projectKey = item.directoryId ? `dir:${item.directoryId}` : sigPart(item.project);
  if (!projectKey) return null;
  return [sigPart(item.client), projectKey, sigPart(item.unitCode), sigPart(item.itemType), sigPart(item.depotArrivalDate), sigPart(item.invoiceNumber)].join("|");
}
// Two consignments of one job, arriving the same day under the same reference, look alike in
// every field the signature covers - and they are not duplicates if they hold different
// cases. The 2501009 job came in twice, twenty-one cases on one entry and one on the other,
// and was offered up for deletion; taking that offer would have thrown away real stock.
//
// So a signature group is split by what it holds. Entries sharing no case are separate
// consignments and are left alone. Entries sharing at least one are still shown together:
// that is the shape of a genuine double import, including one that was edited afterwards.
// An entry with no cases at all can only be judged on its figures, so those stay grouped.
function splitGroupByCases(group) {
  const codesOf = (it) => new Set((it.packages || []).map((p) => String(p.code || "").toUpperCase().replace(/[\s()#'\u2018\u2019]/g, "")).filter(Boolean));
  const buckets = [];
  for (const it of group) {
    const codes = codesOf(it);
    // Nothing to compare on: judged by its figures alongside anything else in that state.
    const hit = codes.size === 0
      ? buckets.find((b) => b.codes.size === 0)
      : buckets.find((b) => [...codes].some((c) => b.codes.has(c)));
    if (hit) {
      hit.items.push(it);
      codes.forEach((c) => hit.codes.add(c));
    } else {
      buckets.push({ items: [it], codes });
    }
  }
  return buckets.map((b) => b.items);
}
// Every check-in the depot has taken, grouped by the reference it came in under. A
// reference states its own package count - "13-DM-25-0625_44PKGS" - so where the entries
// filed against it hold more than that, the same consignment has been checked in twice:
// once as a whole lot and once split per lift. It is invisible on the inventory screen,
// which shows the entries but never adds them back up against the paper.
// A client's storage ledger - "Mitsubishi Inventory Store List" and its equivalents for the
// other makers - is one sheet per site, with a row for every movement in or out and monthly
// billing rows in between that simply repeat the running balance. Only the movement rows
// carry a job number in the first column, which is what tells the two apart; counting the
// billing rows as movements would multiply everything by the number of months it has sat.
//
// The job number is also the only reliable key. An IN row's DM number is the arrival's, but
// an OUT row's is the delivery's own DM - "2512137-OUT | 13-DM-25-0616" - so matching lots
// by DM across the two directions would pair a delivery with whatever lot happened to share
// its number. The FC job number means the same thing on both sides.
function parseClientStoreList(wb) {
  const moves = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: false });
    for (const r of rows) {
      const first = String((r || [])[0] || "").trim();
      const m = first.match(/^(\d{6,8})\s*-\s*(IN|OUT)\s*$/i);
      if (!m) continue;
      const pkgs = Number(String((r || [])[3] || "").replace(/,/g, ""));
      if (!isFinite(pkgs) || pkgs === 0) continue;
      moves.push({
        site: name,
        jobNumber: m[1],
        direction: m[2].toUpperCase(),
        dm: String((r || [])[1] || "").replace(/\s+/g, " ").trim(),
        pkgs: Math.abs(pkgs),
        when: String((r || [])[2] || "").trim(),
      });
    }
  }
  return moves;
}
// Each movement set against what the depot recorded for that same job number. An IN is
// compared with the cases checked in under it, an OUT with the cases delivered under it.
function reconcileStoreList(moves, items) {
  const arrived = {};
  const delivered = {};
  (items || []).forEach((it) => {
    if (it.cancelled) return;
    const j = String(it.jobNumber || "").trim();
    if (j) arrived[j] = (arrived[j] || 0) + totalUnits(it);
    activeDeliveries(it).forEach((d) => {
      const dj = String(d.jobNumber || "").trim();
      if (!dj) return;
      delivered[dj] = (delivered[dj] || 0) + ((d.codes || []).length || 0);
    });
  });
  const bySite = {};
  moves.forEach((mv) => {
    const g = (bySite[mv.site] = bySite[mv.site] || { site: mv.site, rows: [], inSheet: 0, outSheet: 0, inApp: 0, outApp: 0 });
    const app = mv.direction === "IN" ? (arrived[mv.jobNumber] || 0) : (delivered[mv.jobNumber] || 0);
    g.rows.push({ ...mv, app, diff: app - mv.pkgs });
    if (mv.direction === "IN") { g.inSheet += mv.pkgs; g.inApp += app; }
    else { g.outSheet += mv.pkgs; g.outApp += app; }
  });
  return Object.values(bySite)
    .map((g) => ({ ...g, leftSheet: g.inSheet - g.outSheet, leftApp: g.inApp - g.outApp,
      rows: g.rows.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff) || String(a.jobNumber).localeCompare(String(b.jobNumber))) }))
    .sort((a, b) => Math.abs(b.leftApp - b.leftSheet) - Math.abs(a.leftApp - a.leftSheet));
}
// The depot's own movements laid out the way a client's storage ledger lays them out: one
// site at a time, every arrival and every delivery in date order, with a running balance
// down the page. The client's file is the same thing kept by hand, so putting the two side
// by side stops being an act of translation.
//
// Free days and charges are deliberately not here. Those belong to the billing rules and to
// each client's agreement, and a ledger that guessed at them would be worse than one that
// shows only what actually moved.
function storageLedger(items, directory) {
  const sites = {};
  (items || []).forEach((it) => {
    if (it.cancelled) return;
    const key = it.directoryId || sigPart(it.project) || "unspecified";
    const entry = (directory || []).find((d) => d.id === it.directoryId);
    const g = (sites[key] = sites[key] || {
      key, label: (entry && entry.name) || it.project || "\u2014",
      labelZh: (entry && entry.nameZh) || "", client: it.client || "", moves: [],
    });
    const ref = String(it.invoiceNumber || "").trim() || String(it.unitCode || "").trim();
    const arrivals = activeArrivals(it);
    if (arrivals.length) {
      arrivals.forEach((a) => g.moves.push({
        arrivalId: a.id || "", deliveryId: "",
        date: a.date || it.depotArrivalDate || "", dir: "IN",
        jobNumber: it.jobNumber || "", ref, lot: it.unitCode || "",
        pkgs: (a.codes || []).length || totalUnits(it),
        type: a.type || it.arrivingType || "", entryId: it.id,
      }));
    } else {
      g.moves.push({
        arrivalId: "", deliveryId: "",
        date: it.depotArrivalDate || "", dir: "IN", jobNumber: it.jobNumber || "", ref,
        lot: it.unitCode || "", pkgs: totalUnits(it), type: it.arrivingType || "", entryId: it.id,
      });
    }
    activeDeliveries(it).forEach((d) => g.moves.push({
      arrivalId: "", deliveryId: d.id || "",
      date: d.date || "", dir: "OUT", jobNumber: d.jobNumber || "", ref, lot: it.unitCode || "",
      pkgs: (d.codes || []).length,
      // A delivery whose cases all came back still shows, carrying nothing. The trip
      // happened and is charged at the minimum, so dropping the row would hide a movement
      // that has to be invoiced.
      type: d.allReturned ? "Delivery (all returned)" : "Delivery",
      entryId: it.id,
    }));
  });
  return Object.values(sites).map((g) => {
    // Undated movements sort to the end rather than to 1970, where they would drag the
    // running balance negative from the first line.
    const moves = g.moves.slice().sort((a, b) => (a.date ? 0 : 1) - (b.date ? 0 : 1)
      || String(a.date).localeCompare(String(b.date))
      || String(a.jobNumber).localeCompare(String(b.jobNumber)));
    let bal = 0;
    moves.forEach((m) => { bal += m.dir === "IN" ? m.pkgs : -m.pkgs; m.balance = bal; });
    return { ...g, moves, balance: bal, inTotal: moves.filter((m) => m.dir === "IN").reduce((n, m) => n + m.pkgs, 0),
      outTotal: moves.filter((m) => m.dir === "OUT").reduce((n, m) => n + m.pkgs, 0) };
  }).filter((g) => g.moves.length).sort((a, b) => b.balance - a.balance);
}
function checkInAudit(items) {
  const groups = new Map();
  (items || []).forEach((it) => {
    if (it.cancelled) return;
    const ref = String(it.invoiceNumber || "").trim() || "\u2014";
    if (!groups.has(ref)) {
      const m = ref.match(/(\d+)\s*(?:PKGS?|CARTONS?|CASES?)/i);
      groups.set(ref, { ref, stated: m ? Number(m[1]) : null, rows: [], held: 0, remaining: 0 });
    }
    const g = groups.get(ref);
    const units = totalUnits(it);
    const left = Math.max(0, units - deliveredUnits(it));
    g.held += units;
    g.remaining += left;
    // One row per arrival batch, since a batch is what gets reversed. An entry carrying no
    // batches is shown whole, as there is nothing finer to act on.
    const arrivals = activeArrivals(it);
    if (arrivals.length) {
      arrivals.forEach((a) => g.rows.push({
        item: it, arrival: a, units: (a.codes || []).length || units, left,
        date: a.date || it.depotArrivalDate, source: a.declaredSource || it.notes || "",
      }));
    } else {
      g.rows.push({ item: it, arrival: null, units, left, date: it.depotArrivalDate, source: it.notes || "" });
    }
  });
  // What is wrong comes first, then what can be checked, then the rest. Sorting by size put
  // the entries carrying no reference at all - the largest bucket in the depot by far - on
  // top of everything, and the three over-checked references nobody could then find.
  const rank = (g) => (g.over > 0 ? 0 : g.stated != null ? 1 : 2);
  return [...groups.values()]
    .map((g) => ({ ...g, over: g.stated == null ? null : g.held - g.stated }))
    .sort((a, b) => rank(a) - rank(b) || (b.over || 0) - (a.over || 0) || b.held - a.held);
}
function findDuplicateGroups(items) {
  const map = {};
  items.forEach((it) => {
    const sig = itemSignature(it);
    if (!sig) return;
    (map[sig] = map[sig] || []).push(it);
  });
  return Object.values(map)
    .filter((g) => g.length > 1)
    .flatMap(splitGroupByCases)
    .filter((g) => g.length > 1);
}

function todayStr() {
  return dateToLocalISO(new Date());
}
function daysBetween(a, b) {
  if (!a || !b) return null;
  const d1 = new Date(a + "T00:00:00");
  const d2 = new Date(b + "T00:00:00");
  return Math.round((d2 - d1) / 86400000);
}
function addDays(dateStr, n) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() + n);
  return dateToLocalISO(d);
}
function fmt(dateStr) {
  if (!dateStr) return "—";
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function totalUnits(item) {
  if (item.packages && item.packages.length > 0) return item.packages.length;
  const n = Number(item.packageCount);
  return n > 0 ? n : 1;
}
function activeDeliveries(item) {
  return (item.deliveries || []).filter((d) => !d.cancelled);
}
function deliveredUnits(item) {
  if (item.packages && item.packages.length > 0) {
    return activeDeliveries(item).reduce((s, d) => s + (d.codes ? d.codes.length : 0), 0);
  }
  return activeDeliveries(item).reduce((s, d) => s + (Number(d.packageCount) || 0), 0);
}
function remainingUnits(item) {
  return Math.max(0, totalUnits(item) - deliveredUnits(item));
}
function deliveredCodes(item) {
  return activeDeliveries(item).flatMap((d) => d.codes || []);
}
function remainingPackages(item) {
  if (!item.packages || item.packages.length === 0) return [];
  const done = new Set(deliveredCodes(item));
  return item.packages.filter((p) => !done.has(p.code));
}
function activeArrivals(item) {
  return item.arrivals || [];
}
// Cases that can actually leave on a delivery: not yet delivered, and already checked in
// at the depot. An item tracked by arrival batches can carry cases that are on the packing
// list but have not physically landed, and those must not be offered as deliverable - the
// same rule the Delivery screen already applies.
function deliverablePackages(item) {
  const pending = new Set(notYetArrivedPackages(item).map((p) => p.code));
  return remainingPackages(item).filter((p) => !pending.has(p.code));
}
function arrivedCodesSet(item) {
  return new Set(activeArrivals(item).flatMap((a) => a.codes || []));
}
function usesArrivalBatches(item) {
  return activeArrivals(item).length > 0 && (item.packages || []).length > 0;
}
function notYetArrivedPackages(item) {
  if (!usesArrivalBatches(item)) return [];
  const done = arrivedCodesSet(item);
  return item.packages.filter((p) => !done.has(p.code));
}
function earliestArrivalDate(item) {
  const ds = activeArrivals(item).map((a) => a.date).filter(Boolean).sort();
  return ds.length ? ds[0] : "";
}
function effectiveDepotArrivalDate(item) {
  return item.depotArrivalDate || earliestArrivalDate(item);
}
function lastDeliveryDate(item) {
  const ds = activeDeliveries(item).map((d) => d.date).filter(Boolean).sort();
  return ds.length ? ds[ds.length - 1] : null;
}
function remainingShare(item) {
  const total = totalUnits(item) || 1;
  return remainingUnits(item) / total;
}
function remainingWeightKg(item) {
  if (!item.weightKg) return 0;
  return Number(item.weightKg) * remainingShare(item);
}
function remainingVolumeCbm(item) {
  if (!item.volumeCbm) return 0;
  return Number(item.volumeCbm) * remainingShare(item);
}
function depotRemainingTotals(items, depotValue) {
  let kg = 0, cbm = 0, count = 0;
  items.forEach((it) => {
    if (it.depot !== depotValue) return;
    const status = deriveStatus(it);
    if (status !== "at_depot" && status !== "partial") return;
    kg += remainingWeightKg(it);
    cbm += remainingVolumeCbm(it);
    count += 1;
  });
  return { kg: Math.round(kg * 10) / 10, cbm: Math.round(cbm * 1000) / 1000, count };
}

function deriveStatus(item) {
  if (!effectiveDepotArrivalDate(item)) return "pending_collection";
  const hasDeliveries = activeDeliveries(item).length > 0;
  const remaining = remainingUnits(item);
  if (hasDeliveries && remaining <= 0) return "delivered";
  if (hasDeliveries && remaining > 0) return "partial";
  return "at_depot";
}

function lfdAlert(item) {
  if (deriveStatus(item) !== "pending_collection" || !item.terminalLFD) return null;
  const d = daysBetween(todayStr(), item.terminalLFD);
  if (d === null) return null;
  if (d < 0) return { level: "overdue", days: d };
  if (d <= 3) return { level: "soon", days: d };
  return null;
}

function storageInfo(item) {
  const status = deriveStatus(item);
  if (status === "pending_collection") return null;
  const freeDays = freeDaysFor(item);
  const today = todayStr();

  if (usesArrivalBatches(item)) {
    // Each batch's cases run their own free-storage clock from that batch's devan/CFS date.
    const deliveredAt = {};
    activeDeliveries(item).forEach((d) => (d.codes || []).forEach((c) => { deliveredAt[c] = d.date; }));
    let maxBillable = 0;
    let maxHeld = 0;
    let minLeft = Infinity;
    let anyHolding = false;
    activeArrivals(item).forEach((a) => {
      if (!a.date) return;
      (a.codes || []).forEach((code) => {
        const end = deliveredAt[code] || today;
        const held = daysBetween(a.date, end) || 0;
        const over = held - freeDays;
        if (over > 0) maxBillable = Math.max(maxBillable, over);
        if (held > maxHeld) maxHeld = held;
        if (!deliveredAt[code]) {
          anyHolding = true;
          const left = freeDays - held;
          if (left < minLeft) minLeft = left;
        }
      });
    });
    const freeUntil = addDays(earliestArrivalDate(item), freeDays);
    return {
      freeUntil,
      daysHeld: maxHeld,
      billableDays: maxBillable,
      billable: maxBillable > 0,
      freeDays,
      daysLeft: anyHolding && minLeft !== Infinity ? minLeft : null,
    };
  }

  const arrivalDate = effectiveDepotArrivalDate(item);
  const freeUntil = addDays(arrivalDate, freeDays);
  const endDate = status === "delivered" ? lastDeliveryDate(item) || today : today;
  const daysHeld = daysBetween(arrivalDate, endDate);
  const overFree = daysBetween(freeUntil, endDate);
  const billableDays = overFree && overFree > 0 ? overFree : 0;
  return { freeUntil, daysHeld, billableDays, billable: billableDays > 0, freeDays, daysLeft: freeDays - daysHeld };
}

function currentYyMm() {
  const now = new Date();
  return String(now.getFullYear()).slice(-2) + String(now.getMonth() + 1).padStart(2, "0");
}
function allJobNumbersUsed(items) {
  const nums = [];
  (items || []).forEach((it) => {
    if (it.jobNumber) nums.push(it.jobNumber);
    (it.deliveries || []).forEach((d) => {
      if (d.jobNumber) nums.push(d.jobNumber);
    });
  });
  return nums;
}
function nextJobNumber(items) {
  const prefix = currentYyMm();
  let maxSeq = 0;
  allJobNumbersUsed(items).forEach((jn) => {
    if (jn.startsWith(prefix) && jn.length === 7) {
      const seq = Number(jn.slice(4));
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  return `${prefix}${String(maxSeq + 1).padStart(3, "0")}`;
}

function emptyForm() {
  return {
    client: CLIENTS[0],
    project: "",
    invoiceNumber: "",
    itemType: ITEM_TYPES[0],
    packageCount: "",
    unitCode: "",
    constructionSite: "",
    description: "",
    weightKg: "",
    volumeCbm: "",
    // Kept alongside weightKg so the manufacturer's per-case figures survive even after
    // the Devan/CFS sheet total takes over as the billing weight.
    weightPackingListKg: "",
    volumeCbmPackingList: "",
    weightSource: "",
    volumeSource: "",
    oversizeCases: [],
    shkNumber: "",
    ssDoNo: "",
    containers20: "",
    containers40: "",
    zone: "urban",
    cargoType: "",
    arrivingType: ARRIVING_TYPES[0],
    terminalArrivalDate: "",
    terminalLFD: "",
    confirmedCollectionDate: "",
    depot: DEPOTS[0],
    depotArrivalDate: "",
    depotLocation: "",
    plannedDeliveryDate: "",
    deliveries: [],
    arrivals: [],
    packages: [],
    jobNumber: "",
    orderedBy: "",
    poNumber: "",
    jobRef: "",
    directoryId: "",
    recordedBy: "",
    notes: "",
  };
}
const TEXT = {
  en: {
    appSubtitle: "Depot & Storage Manifest",
    navDashboard: "Dashboard",
    navInventory: "Inventory",
    navNewEntry: "New Entry",
    navDeliveries: "Deliveries",
    navDuplicates: "Duplicates",
    navDuplicatesCount: (n) => `Duplicates (${n})`,
    navImport: "Import",
    themeToggleLabel: "Toggle dark mode",
    langToggleLabel: "中文",
    loadingMsg: "Loading depot records…",
    saveErrorMsg: "Could not save — please try again.",
    conflictTitle: "Someone else saved first",
    conflictBody: "Another user saved changes while this page was open, so your last change was not written \u2014 it would have overwritten theirs. The list below now shows what is actually saved. Please redo your change.",
    conflictDismiss: "Got it",

    statAtDepot: "At Depot",
    statPending: "Pending Collection",
    statBillable: "Billable Now",
    statLfd: "LFD Warnings",
    dupBanner: (n) => `${n} possible duplicate ${n === 1 ? "entry" : "entries"} found — same client, project, unit and depot arrival date.`,
    reviewDuplicatesBtn: "Review Duplicates",
    lfdSectionTitle: "Terminal Last Free Day — Action Needed",
    billableSectionTitle: "Currently Billable Storage",
    billableEmptyMsg: "Nothing past the 14-day free period right now.",
    sinceLabel: "since",

    searchLabel: "Search",
    searchPlaceholder: "Project, client, ID, SHK no., case no., job no.",
    clientLabel: "Client",
    statusLabel: "Status",
    statusAll: "All",
    statusPending: "Pending Collection",
    statusAtDepot: "At Depot",
    statusPartial: "Partially Delivered",
    statusDelivered: "Delivered",
    depotLabel: "Depot",
    exportBtn: (n) => `Export to Excel (${n})`,
    newEntryBtn: "+ New Entry",
    colId: "ID",
    colClient: "Client",
    colProjectSite: "Project / Site",
    colUnit: "Unit",
    deliverySearchPlaceholder: "Client, project, unit no., job no., or case no.",
    colDepot: "Depot",
    colDepotArrival: "Depot Arrival",
    colStatus: "Status",
    noRecordsMsg: "No records match. Add a new entry to get started.",
    editBtn: "Edit",
    deleteBtn: "Delete",
    deliverBtn: "Deliver",
    duplicateBadge: "DUPLICATE?",

    titleNew: "New Manifest Entry",
    titleEdit: "Edit Manifest Entry",
    fClient: "Client",
    fProject: "Project",
    fProjectEn: "Project / Site (English)",
    fProjectZh: "Project / Site (Chinese)",
    fProjectPlaceholder: "e.g. MTR Yau Tong Station MOD",
    fInvoiceNo: "Invoice No.",
    fInvoiceHint: "Fill in once invoiced",
    sectionSite: "Construction Site",
    fSiteName: "Construction Site Name",
    fSiteHint: "English or Traditional Chinese — one is fine, no need for both",
    fSitePlaceholder: "e.g. 油塘站 or MTR Yau Tong Station",
    sectionCargo: "Cargo Detail",
    fItemType: "Item Type",
    fPackages: "# of Packages",
    fUnitCode: "Escalator / Elevator # or Code",
    fUnitCodeHint: "e.g. L1, E02",
    fDescription: "Description",
    fDescriptionHint: "e.g. unit no. / equipment description",
    fReference: "Reference / SHK No.",
    fReferenceHint: "Schindler reference, if applicable",
    fWeight: "Weight (KG)",
    fVolume: "Volume (CBM)",
    f20: "No. of 20' Containers",
    fZone: "Zone",
    fZoneHint: "Affects the Devan/Delivery handling rate for clients with a tariff on file",
    zoneUrban: "Urban Area",
    zoneLantau: "Lantau / Tung Chung",
    fCargoType: "Cargo Type",
    fCargoTypeHint: "Leave on Auto to detect from the description",
    cargoTypeAuto: "Auto-detect",
    cargoTypeElevator: "Elevator",
    cargoTypeEscalator: "Escalator",
    f40: "No. of 40' Containers",
    sectionArrival: "Arrival & Depot",
    fArrivingType: "Arriving Type",
    fArrivingTypeHint: "Devan = we unpack a container ourselves. CFS = client delivers loose packages, no devanning needed.",
    fDepot: "Depot",
    fDepotHint: "Which depot this is going into",
    fDepotLocation: "Depot Location / Bay",
    fTerminalArrival: "Terminal Arrival Date",
    fTerminalArrivalHint: "ETA at container terminal",
    fTerminalLFD: "Terminal Last Free Day",
    fTerminalLFDHint: "Latest date to collect before demurrage",
    fConfirmedCollection: "Confirmed Collection Date",
    fConfirmedCollectionHint: "Date we go to terminal",
    fDepotArrival: "Depot Arrival Date",
    fDepotArrivalHint: "Starts the 14-day free storage clock",
    fPlannedDelivery: "Planned Delivery Date",
    fPlannedDeliveryHint: "Estimate only — record real deliveries via the Deliveries tab",
    deliveryProgress: (del, tot, count) => `${del} of ${tot} unit(s) delivered across ${count} delivery record(s), last on `,
    deliveryProgressManage: "Manage deliveries via the Deliveries tab.",
    fNotes: "Notes",
    saveBtn: "Save Entry",
    cancelBtn: "Cancel",

    deliveryTitlePrefix: "Record Delivery —",
    combinedDeliveryTitle: (n) => `Record Combined Delivery — ${n} entries`,
    combinedHistoryHiddenNote: "Past deliveries for these entries aren't shown here — open each entry individually (from Inventory) to view or cancel its delivery history.",
    recordCombinedBtn: (n) => n > 0 ? `Record Delivery for Selected (${n})` : "Record Delivery for Selected",
    combinedSelectionHint: "Tick every entry going out together on the same job \u2014 e.g. two different batches heading to the same site on the same day \u2014 then record them as one combined delivery with a single job sheet.",
    selectedTag: "Included \u2713",
    includeInDeliveryBtn: "Include in this delivery",
    addCombinedDeliveryBtn: "Record Combined Delivery",
    combinedPrintLabel: "Combined Delivery Job Sheet",
    addMoreBatchesBtn: "+ Add more batches to this delivery",
    addMoreBatchesSearchPlaceholder: "Search by unit no. or job no.",
    addMoreBatchesScopeNote: (site) => `Showing other entries at ${site} only.`,
    addMoreBatchesNoneMsg: "No other entries at this same construction site are available to add.",
    plannedWasText: (date) => ` · planned delivery was ${date}`,
    progressOf: "of",
    progressDeliveredSoFar: "unit(s) delivered so far",
    progressRemaining: "remaining at the depot",
    colDate: "Date",
    colQty: "Qty",
    colDeliveredTo: "Delivered To",
    colReceivedBy: "Received By",
    removeBtn: "Remove",
    fDeliveryDate: "Delivery Date",
    fDeliveryDateHint: "This is what closes the storage clock once fully delivered",
    fQty: "Quantity Delivered",
    fQtyHint: (r) => `${r} remaining`,
    fDeliveredTo: "Delivered To",
    fDeliveredToHint: "Defaults to the construction site",
    fReceivedBy: "Received By",
    fReceivedByHint: "If left blank, uses the Ordered By contact for this project",
    overshootMsg: (r) => `Only ${r} unit(s) remain at the depot for this entry — reduce the quantity.`,
    addDeliveryBtn: "Add Delivery Record",
    closeBtn: "Close",
    allDeliveredMsg: "All units for this entry have been delivered.",
    selectItemMsg: "Select an item to record a delivery. Items with more than one package (e.g. a lot of escalator crates) can be delivered in stages — each stage reduces what's left at the depot.",
    nothingAtDepotMsg: "Nothing is currently at the depot.",
    recordDeliveryBtn: "Record Delivery",

    noneFoundMsg: "No duplicates found. Entries are flagged when client, project, unit code, item type, depot arrival date, and invoice number all match another entry.",
    matchingEntries: (n) => `${n} matching entries`,
    dupSharedLine: (arrival, type, jobs) => `Same arrival ${arrival} \u00b7 ${type} \u00b7 job no. ${jobs}`,
    dupColCases: "Cases",
    dupColDeliveries: "Deliveries",
    dupFlatCount: (n) => `${n} (no case list)`,
    dupSourceLabel: "From",
    dupHasDeliveriesWarn: (n) => `${n} delivery record${n === 1 ? "" : "s"} on this entry \u2014 deleting it removes them too`,
    dupDiffHint: "Highlighted values are the ones that differ between these entries. Everything else already matches, which is why they were flagged.",
    deleteAllBtn: (n) => `Delete all ${n}`,
    colInvoiceNo: "Invoice No.",
    colAddedOn: "Added On",
    keepDeleteBtn: "Keep this, delete others",

    tabExcel: "Excel Upload",
    tabPdf: "PDF Scan",
    pdfCaseCountMismatch: (lot, stated, read) => `${lot}: the document states ${stated} package${stated === 1 ? "" : "s"} but ${read} case number${read === 1 ? "" : "s"} were read, ${Math.abs(stated - read)} ${stated > read ? "short" : "over"} \u2014 check the C/S NO. list against the paper. The ${read} read are used as they stand; nothing has been invented to make up the difference.`,
    pdfWholeDocument: "This document",
    pdfWeightMismatch: (stated, read) => `The document totals ${stated} kg but the cases read come to ${read} kg \u2014 check the lots below against the paper before importing.`,
    pdfCasesCorrected: (lot, changes) => `${lot}: case numbers corrected to match the document's own Shipping Marks \u2014 ${changes}. Weights and volumes are unchanged.`,
    pdfRepeatedLots: (lots) => `More than one group came back named ${lots}. If the document splits an order into groups going to different lifts, give each its own name before importing, or they will be checked in as one lot.`,
    pdfDocumentTotals: (cbm, kg) => `The document gives one total for the whole shipment \u2014 ${cbm} CBM, ${kg} kg \u2014 without splitting it between the orders, so it has not been divided up. Enter the volume per lot from the paperwork if you have it.`,
    pdfTerminalDatesFound: (eta, lastFree) => `This document also gives terminal dates \u2014 arrival ${eta}, last free day ${lastFree}. Enter them on the entry when you check these cases in; they are not part of the packing list.`,
    tabManualPackingList: "Manual Entry",
    manualLotBuilderLabel: "Build a lot from the job sheet",
    manualLotBuilderHint: "One CFS sheet often names several lots. Fill this row for each and press Add lot \u2014 they collect below and go in together. Type the C/S NO. marking exactly as it appears (1,2,3/3 \u00b7 1-12/23 \u00b7 1/2, 2/2); the cases are made for you and the lot's stated weight and volume split evenly across them. For a lot whose case codes aren't plain numbers, leave C/S No. empty and list them under Itemized Packages instead.",
    manualOrderNoLabel: "Order no.",
    manualCaseSpecLabel: "C/S No.",
    manualAddLotBtn: "Add lot",
    manualCaseSpecPreview: (n, codes) => `${n} case${n === 1 ? "" : "s"}: ${codes}`,
    manualPendingLotsLabel: (n) => `${n} lot${n === 1 ? "" : "s"} ready to add`,
    manualUnnamedLot: "(unnamed lot)",
    legacySameOrderWarn: (orders) => `Order no. ${orders} appears on more than one shipment at this site. An order can ship in batches, and each packing list numbers its own cases, so these may be separate consignments that happen to share case numbers \u2014 or the same one recorded twice. Check the paperwork before choosing where to check these cases in.`,
    legacySameOrderSibling: (id, unit, n) => `${id} \u00b7 ${unit} \u00b7 ${n} case${n === 1 ? "" : "s"} not yet checked in`,
    legacySameOrderOnlyThere: (codes) => `carries ${codes}, which this one does not`,
    manualPackingListOpenBtn: "+ Add Packing List",
    manualPackingListDesc: "For older jobs with no packing list file, or one that's incomplete \u2014 type in the case list by hand. This creates one Incoming shipment the same way an uploaded file would, ready to check in via Devan/CFS.",
    legacyManualEntryNote: "Manually entered \u2014 no packing list file on file for this shipment.",
    excelTitle: "Import from Excel",
    excelDesc: 'Column headers are matched against the depot\'s field names automatically (e.g. "Invoice No.", "Depot Arrival Date"). Unrecognized columns are skipped and listed below.',
    chooseFileBtn: "Choose File (.xlsx, .xls, .csv)",
    downloadTemplateBtn: "Download blank template",
    selectedCount: (sel, tot) => `${sel} of ${tot} selected to import.`,
    selectAllBtn: "Select all",
    clearBtn: "Clear",
    legacyRangeTapHint: "Tap a case, then shift-tap another to take everything between them.",
    selectNonDupBtn: "Select non-duplicates only",
    unmatchedMsg: "Columns not recognized (skipped): ",
    prevColClient: "Client",
    prevColProject: "Project",
    prevColItemType: "Item Type",
    prevColDepot: "Depot",
    prevColDepotArrival: "Depot Arrival",
    prevColMatch: "Match",
    importBtn: (n) => `Import ${n} Record(s)`,
    discardBtn: "Discard",
    pdfTitle: "Scan a PDF Packing List",
    pdfDesc: "Upload a client's packing list, delivery memo, or shipping list PDF — any layout. It reads the lift/lot breakdown, cases, weights and CBM automatically, the same way the Excel packing list import does.",
    choosePdfBtn: "Choose PDF",
    scanningMsg: "Reading document…",
    pdfReadErrorMsg: "Couldn't automatically read this PDF. Please check the file, or enter the details manually below.",
    pdfTooLargeMsg: (kb, why) => `This PDF (${kb} KB) was too much for the scanner to finish \u2014 long packing lists that list every case's contents often are. Try the Excel version of the list if there is one, split the PDF, or enter the lots by hand below. (${why})`,
    pdfTooLargeMsg: (kb, why) => `此PDF（${kb} KB）內容過多，掃描未能完成 \u2014 逐件列出內容之裝箱單常有此情況。如有Excel版本請改用，或將PDF分割，亦可於下方手動輸入。（${why}）`,
    pdfTruncatedMsg: "This document is very dense (lots of cases/items) and the automatic read got cut off partway through. Try again \u2014 it sometimes succeeds on a second attempt \u2014 or enter the details manually below.",
    pdfKeyWarning: "This uses your own Anthropic API key, entered below and saved only in this browser. Since it's used directly from this page, anyone who opens this file's developer tools could see it — fine for internal testing among trusted staff, not for wider distribution.",
    pdfApiKeyLabel: "Anthropic API Key",
    pdfApiKeyHint: "From console.anthropic.com — saved only in this browser",
    pdfSaveKeyBtn: "Save Key",
    pdfKeySavedBadge: "Key saved",
    pdfNeedKeyMsg: "Add and save an API key above to enable PDF scanning.",
    reviewWarningMsg: "Review the extracted details below before saving — automatic reads can miss or misplace things.",
    dupWarningMsg: (match) => `This looks like it may match an existing entry: ${match.id} (${match.client} · ${match.project}).`,
    excelErrorMsg: "Couldn't read that file. Make sure it's a .xlsx, .xls, or .csv export.",
    excelNoRowsMsg: "No rows found in that file.",

    badgePendingCollection: "PENDING COLLECTION",
    badgeLfdOverdue: (d) => `LFD OVERDUE ${d} DAYS`,
    badgeLfdSoon: (d) => `LFD IN ${d} DAYS`,
    badgeBillable: (d) => `BILLABLE · ${d} DAYS`,
    badgeFree: (d) => `FREE · ${d} DAYS LEFT`,
    badgePartial: (del, tot, extra) => `PARTIAL ${del}/${tot}${extra}`,
    badgeDelivered: (extra) => `DELIVERED${extra}`,
    badgeBilledSuffix: (d) => ` · BILLED ${d} DAYS`,
    badgeBillableSuffix: (d) => ` · ${d} DAYS`,
    badgeDupOf: (id) => `DUP OF ${id}`,
    badgeNew: "NEW",

    sectionPackages: "Itemized Packages",
    packagesHint: "Optional \u2014 add individual case/package codes so specific ones can be delivered separately (e.g. case 2A, 13A). Leave empty to just track a total count instead.",
    bulkAddLabel: "Quick add (comma or line separated)",
    bulkAddPlaceholder: "e.g. 1A, 2A, 3A, 12A, 13A",
    bulkAddBtn: "Add Codes",
    packagesGenerateLabel: "How many packages?",
    packagesGenerateBtn: "Generate Rows",
    packagesTotalLabel: (n) => `Total: ${n} package${n === 1 ? "" : "s"}`,
    packageCodeCol: "Code",
    packageDescCol: "Description",
    packageWeightCol: "Weight (kg)",
    packageCbmCol: "CBM",
    removePackageBtn: "Remove",
    packagesCountSummary: (n) => `${n} itemized package(s)`,
    noPackagesMsg: "No itemized packages yet \u2014 using the plain count above instead.",

    uploadModePackingList: "Packing List (Excel/PDF)",
    uploadModeLegacy: "Legacy Upload (CFS/Devan/Delivery)",
    tabPackingList: "Packing List Import",
    packingListTitle: "Import a Client Packing List",
    packingListDesc: "Upload a client's packing list as-is (TK Elevator, Schindler, OTIS, etc. all use different layouts \u2014 this reads the columns automatically). Each distinct lift/lot found becomes its own manifest entry with its cases itemized.",
    choosePackingListBtn: "Choose Packing List File",
    packingListCaseCountWarn: (list) => `Read, but check these before you add them: ${list}. The packages column and the case list disagree, so the depot would take in a different number of cases than the paperwork declares.`,
    packingListNoStructure: "Couldn't recognize a packing-list table in that file. It may use a layout this hasn't seen before \u2014 try Excel Upload instead, or add entries manually.",
    packingListDetectedTitle: (n) => `Found ${n} lift/lot(s) in this file`,
    packingListCommonFieldsTitle: "Apply to all of these",
    packingListApplyClient: "Client",
    packingListApplyProject: "Project",
    packingListApplyDepot: "Depot",
    packingListApplyDepotArrival: "Depot Arrival Date (all items arrived together)",
    packingListApplyDepotLocation: "Depot Location / Bay",
    packingListImportBtn: (n) => `Import ${n} Manifest Entr${n === 1 ? "y" : "ies"}`,
    colLot: "Lift / Lot",
    colPackages: "Packages",
    colContainers: "Container(s)",
    colWeight: "Weight (kg)",
    selectByBatchLabel: "Select by batch (this delivery can pull from more than one arrival batch)",
    selectCodesLabel: "Select which package codes are going out",
    noCodesRemainingMsg: "All itemized packages for this entry have been delivered.",

    colCbm: "CBM",
    colKg: "KG",
    colJobNo: "Job No.",

    sectionJobSheet: "Job Sheet",
    fJobNumber: "Job No.",
    fJobNumberHint: "YYMM + running number, shared across Devan/CFS/Delivery",
    generateJobNoBtn: "Generate Job No.",
    fOrderedBy: "Ordered By",
    fPoNumber: "P.O. No.",
    fPoNumberHint: "Provided by the client",
    fJobRef: "Job Ref. (Site Code)",
    fJobRefHint: "e.g. KTN-002, GAGE-001",
    printJobSheetBtn: "Print Job Sheet / Save as PDF",
    printBtn: "Print / Save as PDF",
    closePreviewBtn: "Close",

    jsTitle: "JOB SHEET",
    jsTitleZh: "工　單",
    jsFrom: "FROM",
    jsFromZh: "由",
    jsTo: "TO",
    jsToZh: "送",
    jsAccount: "ACCOUNT",
    jsAccountZh: "客戶",
    jsJobNo: "JOB NO.",
    jsJobNoZh: "快達單號",
    jsDate: "DATE",
    jsDateZh: "日期",
    jsOrderedBy: "ORDERED BY",
    jsOrderedByZh: "落單人",
    jsPoNo: "P.O. NO.",
    jsPoNoZh: "採購編號",
    jsJobRef: "JOB REF.",
    jsJobRefZh: "地盤代號",
    jsDescription: "DESCRIPTION",
    jsDescriptionZh: "貨物資料／工作程序",
    jsIssuedBy: "ISSUED BY",
    jsIssuedByZh: "出單人",
    jsTotal: "TOTAL",
    jsTotalZh: "共",
    jsPkgs: "PKGS",
    jsKgs: "KGS",
    jsCbm: "CBM",
    jsDevanFrom: (dep) => `DEVAN AT ${dep}`,
    jsCfsFrom: "CFS \u2014 CLIENT DELIVERED, NO DEVANNING",
    jsDeliveryType: "DELIVERY",
    jsDevanType: "DEVAN",
    jsCfsType: "CFS",
    jsSignatureLine: "Customer signature confirming above work completed:",
    jsTemplateLabel: "Template",
    jsCfsFromPreset: "CFS location…",
    jsEditableHint: "Dashed boxes can be edited before printing.",
    jsOversizeLabel: "Oversize Cases 超長/超大件:",
    jsOversizePlaceholder: "#case codes @X.XXCBM per line, then (合共 Total: X CBM)  xN 倍",
    jsOversizeNote: "Oversize CBM is charged at the listed multiplier. Applies to Schindler & Chevalier accounts only.",
    deliveryHistoryLabel: "Delivery History",
    signedDocArrivalLabel: "Signed arrival form (Devan/CFS)",
    signedDocUploadBtn: "Upload signed copy",
    signedDocReplaceBtn: "Replace signed copy",
    signedDocViewBtn: "View signed copy",
    signedDocSavingMsg: "Saving…",
    signedDocFailMsg: "Couldn't save this file. Photos are compressed automatically, but PDFs must be under about 3 MB — try a photo of the signed sheet instead.",
    fSsDoNo: "SS/D.O. No. (提單資料)",
    fSsDoNoHint: "Vessel, voyage & container numbers — prefills the SS/D.O. row on Devan/CFS job sheets",
    jsEstimatedNote: "~ Weight/CBM estimated as a proportional share of the full entry — not individually weighed per package.",

    navDirectory: "Directory",
    tabSitesAccounts: "Sites & Accounts",
    tabEmployees: "Employees",

    dirTitle: "Site & Account Directory",
    dirDesc: "One entry per real construction site \u2014 keeps client, job ref, and contact consistent, and stops the same site being typed differently (English vs Chinese) on different entries.",
    dirAddBtn: "+ Add Site",
    fSiteEn: "Site Name (English)",
    fSiteZh: "Site Name (Chinese)",
    fDirClient: "Account (Client)",
    fDirJobRef: "Job Ref. (Site Code)",
    fDirOrderedBy: "Default Ordered By",
    fDirOfficer: "Responsible Account Officer",
    dirColSite: "Site",
    dirColClient: "Account",
    dirColJobRef: "Job Ref.",
    dirColOfficer: "Account Officer",
    dirColOrderedBy: "Ordered By",
    dirNoneMsg: "No sites added yet.",
    selectFromDirectory: "Fill from Directory",
    selectFromDirectoryPlaceholder: "\u2014 pick a site to auto-fill \u2014",
    showOlderJobs: "Show older jobs",
    showOlderJobsCount: (n) => ` (${n} finished ${n === 1 ? "site" : "sites"} hidden)`,
    manualLinkedToDirectory: "Linked to this site in the Directory \u2014 its Chinese name, job ref and ordered-by have been filled in.",
    saveNewSiteToDirectory: (name) => `Save "${name}" as a new site in the Directory, so future imports recognize it automatically`,

    siteTotalsTitle: "CBM & KG Remaining by Construction Site",
    siteRefsColRef: "Reference (DM / SHK)",
    siteRefsColLots: "Lifts",
    siteRefsColCases: "Case numbers still at the depot",
    siteTotalsColLastCfs: "Last CFS",
    siteTotalsColLastDevan: "Last Devan",
    siteTotalsColLastReturn: "Last return",
    siteTotalsColLastDelivery: "Last delivery",
    siteTotalsColSite: "Construction Site",
    siteTotalsColClient: "Client",
    siteTotalsColPkgs: "Pkgs Left",
    siteTotalsColCbm: "CBM Left",
    siteTotalsColKg: "KG Left",
    siteTotalsNoneMsg: "Nothing currently at the depot.",
    siteTotalsToggleHide: "Hide",
    siteTotalsToggleShow: "Show",

    empTitle: "Employees",
    empDesc: "Add anyone who does data entry, devans/CFS, or deliveries. Roles can be picked from the list or typed fresh.",
    empAddBtn: "+ Add Employee",
    fEmpName: "Name",
    fEmpRole: "Role",
    fEmpRolePlaceholder: "Pick or type a role",
    empColName: "Name",
    empColRole: "Role",
    empNoneMsg: "No employees added yet.",

    signedInAs: "Signed in as",
    mobileMenuLabel: "Menu",
    loginNameLabel: "Name",
    loginNamePlaceholder: "Select your name",
    loginPasswordLabel: "Password",
    loginErrorMissing: "Enter your name and password.",
    loginErrorWrong: "Wrong name or password.",
    loginBtn: "Log In",
    loginBusyMsg: "Checking\u2026",
    loginDefaultPwHint: "First time? The default password is \"Farspeed\" \u2014 you can change it after logging in.",
    changePasswordLink: "Change Password",
    changePasswordTitle: "Change Password",
    logoutBtn: "Log Out",
    pwCurrentLabel: "Current Password",
    pwNewLabel: "New Password",
    pwConfirmLabel: "Confirm New Password",
    pwTooShortMsg: "New password must be at least 4 characters.",
    pwMismatchMsg: "New password and confirmation don't match.",
    pwCurrentWrongMsg: "Current password is incorrect.",
    pwChangedMsg: "Password changed. Use your new password next time you log in.",
    pwSaveBtn: "Save New Password",
    signedInNone: "Not signed in",

    fRecordedBy: "Recorded By",
    fRecordedByHint: "Required \u2014 who is doing this Devan / CFS / Delivery",
    recordedByRequiredMsg: "Select who is recording this before saving.",
    saveBlockedMsg: (fields) => `Fill in ${fields} before saving.`,

    resetBtn: "Reset All Deliveries (demo)",
    resetConfirmMsg: "This clears every delivery record on every item so everything shows as not yet delivered. It does not delete any manifest entries. Continue?",
    resetDoneMsg: "All delivery records cleared.",

    navJobLog: "Job Log",
    navBilling: "Billing",
    navUpload: "Upload",
    navIncoming: "Incoming",
    incomingTitle: "Incoming",
    incomingDesc: "Cases from uploaded packing lists that haven't been checked into the depot yet. Select which cases arrived via Devan or CFS and when \u2014 this creates or updates the real inventory entry.",
    incomingUploadToggle: "Upload Packing List",
    incomingUploadShow: "Show",
    incomingUploadHide: "Hide",
    incomingShowCompleted: "Show fully checked-in shipments",
    incomingNoneMsg: "Nothing incoming right now \u2014 upload a packing list to add cases here.",
    incomingCaseCount: (n) => `${n} case${n === 1 ? "" : "s"} on packing list`,
    incomingLinkedTo: (id) => `linked to ${id}`,
    incomingFullyCheckedIn: "Fully checked in",
    incomingRemainingBadge: (remaining, total) => `${remaining} of ${total} not yet checked in`,
    incomingSelectCasesLabel: "Select which cases are coming in",
    incomingCheckInBtn: (n) => n > 0 ? `Check In (${n})` : "Check In",
    incomingCheckedInNote: (incId) => `Checked in from Incoming ${incId}`,
    packingListAddToIncomingBtn: (n) => `Add ${n} Group${n === 1 ? "" : "s"} to Incoming`,
    packingListShipmentCbmLabel: "Whole-shipment CBM",
    packingListDistributeCbmBtn: "Split by weight",
    packingListDistributeCbmHint: "Shares one document-wide volume across the lots in proportion to their weight. An estimate \u2014 type real per-lot figures in instead where you have them.",
    dirInlineEditBtn: "Edit this site",
    dirInlineEditTitle: "Directory entry",
    dirInlineEditHint: "Saved to the Directory straight away, so every future import of this site picks it up.",
    dirInlineSavedMsg: "Site updated.",
    packingListColProject: "Project / Site",
    packingListAddSiteBtn: (code) => `Add ${code} to the Directory`,
    packingListSitesLabel: (n) => `Project site${n === 1 ? "" : "s"} on this import (${n})`,
    packingListProjectFromCommon: "\u2014 use the fields above \u2014",
    packingListProjectUnknown: (code) => `${code} is not in the Directory \u2014 pick a site or add it first.`,
    packingListMultiProjectHint: (n) => `This file covers ${n} projects. Each lot goes to its own site, so check the column before importing \u2014 the fields above apply only to lots left unset.`,
    packingListCasesHint: "Case numbers for this lot, comma-separated. Renaming keeps each case's weight and volume; adding or removing one changes what the lot holds, and the totals above follow.",
    packingListRemoveGroupBtn: "Remove this group",
    incomingDeleteBtn: "Delete this Incoming shipment",
    incomingDeleteConfirm: "Delete this Incoming shipment? This only removes it from Incoming \u2014 any inventory entry already checked in from it is not affected.",
    packingListIncomingHint: "This just adds the cases to Incoming \u2014 depot, Devan/CFS, job number, and date get decided later when you check them in from the Incoming tab.",
    billingTitle: "Storage Billing",
    billingDesc: "Search by client, project, job number, or case to see storage charges. Billed per arrival batch (not split by construction site) at the client's $/CBM monthly rate: free days first, the remainder of that month pro-rated by day, then every following month in full — even if the goods leave partway through it.",
    billingSearchPlaceholder: "Client, project, job no., or case no.",
    billingColClient: "Client",
    billingColProject: "Project / Site",
    billingColJobNo: "Job No.",
    billingColBatchDate: "Arrival",
    billingColCbm: "CBM",
    billingColRate: "Rate",
    billingColStatus: "Status",
    billingColTotal: "Total",
    billingNoneMsg: "No billable storage found for this search.",
    billingPerCbmMonth: "CBM/mo",
    billingOngoing: "ONGOING",
    billingClosed: "DELIVERED",
    billingShowBtn: "Show breakdown",
    billingHideBtn: "Hide",
    billingCasesLabel: "Cases",
    billingFreeDaysNote: (d) => `${d} free days applied before billing starts.`,
    billingEstimatedNote: "* CBM split estimated from package count share (no per-case CBM available for the delivered/remaining split).",
    billingGrandTotal: "Grand Total",
    billingFootnote: "Ongoing rows are calculated up to today and will keep growing until the goods are marked delivered. Rates are set in Directory → CBM Pricing.",
    billingSelectAllAria: "Select all entries shown",
    billingSelectedCount: (entries, rows, total) => `${entries} entr${entries === 1 ? "y" : "ies"} selected \u2014 ${rows} billing row${rows === 1 ? "" : "s"}, ${total}`,
    billingDeleteSelectedBtn: (n) => `Delete ${n} entr${n === 1 ? "y" : "ies"}`,
    billingDeleteSelectedHint: "One password for the whole batch. Deleting an entry removes every billing row behind it, and its arrival and delivery records with it.",
    billingDeleteItemBtn: "Delete this entry (admin password required)",
    adminConfirmTitle: "Confirm Admin Action",
    adminConfirmDesc: (name) => `Re-enter ${name || "your"}'s password to continue. This permanently deletes the underlying inventory entry, not just this billing row.`,
    adminConfirmBtn: "Delete Permanently",
    billingModeSearch: "Search",
    billingModeMonthly: "Monthly Summary",
    billingModeHandling: "Handling Charges",
    billingHandlingDesc: "Devan/Delivery handling fees and container haulage, per the client tariff \u2014 separate from storage billing above. Only clients with a tariff on file (Chevalier, Schindler) are shown.",
    billingHandlingNeedsQuote: (n) => `${n} ${n === 1 ? "job needs" : "jobs need"} a manual quote \u2014 the tariff doesn't give an automatic rate for that zone/cargo combination.`,
    billingHandlingNoneMsg: "No handling charges yet \u2014 these appear once an item with a tariff-covered client has an arrival or delivery recorded.",
    billingHandlingColType: "Job",
    billingHandlingColBasis: "Basis",
    billingHandlingColRate: "Rate",
    billingHandlingTypeDevan: "Devan",
    billingHandlingTypeDelivery: "Delivery",
    billingHandlingTypeHaulage: "Container Haulage",
    billingHandlingOversizeTag: (mult) => `oversize \u00d7${mult}`,
    billingHandlingHaulageBasis: (c20, c40) => [c20 ? `${c20} \u00d7 20'` : "", c40 ? `${c40} \u00d7 40'` : ""].filter(Boolean).join(", "),
    billingHandlingQuoteBadge: "Quote separately",
    billingHandlingFootnote: "R/Ton = revenue ton = max(weight in tons, volume in CBM). Rates from the 2018 tariff sheets; zone and cargo type are set per item (edit in Manual Entry). Oversize multiplier reuses the same tiers shown on the printed job sheet.",
    billingModeInvoices: "Invoices",
    invoicesTitle: "Debit Notes \u2014 Checked Against Our Records",
    invoicesDesc: "Record each debit note sent to a client, and it is set against what this app's own records produce for the same client, site, category and month. Anything that doesn't agree is flagged before the client finds it.",
    invoicesNoneMsg: "No debit notes recorded for this month.",
    invoicesFootnote: "Covers Storage, CFS, Devan and Delivery. Hoisting and Shifting are not billed from here yet. Click a row to see the charges behind our figure.",
    invoiceAddBtn: "Add Debit Note",
    invoiceColNo: "Debit Note No.",
    invoiceColDate: "Date",
    invoiceColSite: "Site",
    invoiceColCategory: "Category",
    invoiceColInvoiced: "Invoiced",
    invoiceColExpected: "Our records",
    invoiceColDifference: "Difference",
    invoiceStatusMatch: "Agrees",
    invoiceStatusOver: "Overcharged",
    invoiceStatusUnder: "Undercharged",
    invoiceStatusNothing: "No charges found",
    invoiceNothingHint: "Our records show nothing chargeable for this client, site, category and month. Check the site name matches, and that the arrivals or deliveries were recorded.",
    invoiceLinesLabel: (n) => `${n} charge${n === 1 ? "" : "s"} behind our figure`,
    invoiceEstimatedTag: "estimated",
    invoiceTotalInvoiced: "Total invoiced",
    invoiceTotalExpected: "Total per our records",
    invoiceTotalDifference: "Difference",
    invoiceProblemCount: (n) => `${n} note${n === 1 ? "" : "s"} to check`,
    invoiceUninvoicedLabel: "Charges this month with no debit note recorded",
    invoiceViewScanBtn: "View scan",
    invoiceDueDateLabel: "Payment due",
    invoiceAmountLabel: "Amount (HK$)",
    invoiceOrderedByLabel: "Ordered by",
    invoiceOrderRefLabel: "Order reference",
    invoiceChargeLineLabel: "Charge line",
    invoiceChargeLineHint: "The wording that appears on the note, e.g. CHARGES FOR JANUARY 2026 AT LUMPSUM",
    invoiceNarrativeLabel: "Extra line (optional)",
    invoiceBillToLabel: "Bill to address",
    invoiceRevisedDateLabel: "Revised on",
    invoiceRevisedByLabel: "Revised by",
    invoiceScanLabel: "Attach scan",
    invoiceAttachingMsg: "Attaching scan\u2026",
    invoicePreviewMsg: (expected, diff) => `Our records give ${expected} for this client, site, category and month \u2014 difference ${diff}.`,
    debitNotePrintLabel: (no) => `Debit Note ${no}`,
    billingMonthLabel: "Month",
    billingYearLabel: "Year",
    billingMonthNoneMsg: "No storage charges fall in this month.",
    billingMonthFootnote: "Each client's total is what should match their MYOB invoice for this month — use this to double-check before billing. Click a client to see every line item behind their total.",
    jobLogSearchPlaceholder: "Job no., site, job ref, lift, SHK\u2026",
    jobLogFromLabel: "From",
    jobLogToLabel: "To",
    jobLogNoMatchMsg: "No job sheets match these filters.",
    jobLogCount: (shown, total, jobs) => shown === total
      ? `${total} sheet${total === 1 ? "" : "s"} \u00b7 ${jobs} job number${jobs === 1 ? "" : "s"}`
      : `${shown} of ${total} sheets \u00b7 ${jobs} job number${jobs === 1 ? "" : "s"}`,
    jobLogTitle: "All Job Numbers Used",
    jobLogDesc: "Every Devan, CFS, and Delivery job number ever created, most recent first. Click any row to view and reprint that job sheet.",
    jobLogColJobNo: "Job No.",
    jobLogColType: "Type",
    jobLogColDate: "Date",
    jobLogColClient: "Client",
    jobLogColSite: "Project / Site",
    jobLogColRecordedBy: "Recorded By",
    jobLogNoneMsg: "No job numbers created yet.",
    viewReprintBtn: "View / Reprint",

    navCancelledJobs: "Cancelled Jobs",
    cancelJobBtn: "Cancel",
    cancelledJobsTitle: "Cancelled Jobs",
    cancelledJobsDesc: "Voided Devan, CFS, and Delivery job sheets. They're kept here (not the main Job Log) so the job number stays reserved. Since this is a demo, you can permanently delete them from here to clean up test data.",
    cancelledJobsNoneMsg: "Nothing cancelled.",
    restoreBtn: "Restore",
    purgeBtn: "Permanently Delete",
    permanentDeleteConfirmMsg: "This permanently deletes the record and frees up its job number. This cannot be undone. Continue?",

    inventoryRemainingLabel: "remaining in depot",
    inventoryNoRemainingPkgsMsg: "All itemized packages have left the depot.",

    settingsLabel: "Settings",
    navDuplicatesShort: "Duplicates",
    navCheckIns: "Check-ins",
    navCheckInsCount: (n) => `Check-ins (${n} over)`,
    fAwaitingCollection: "No arrival date yet \u2014 these goods are still awaiting collection",
    legacyFieldNames: { client: "client", site: "site", date: "date" },
    legacyRowMissingHint: (list) => `This file needs its ${list} filled in before it can be processed. The date decides where it sits in the storage ledger and when storage starts being charged, so it cannot be left blank.`,
    navPlReader: "Packing list reader",
    plrTitle: "Packing list reader",
    plrDesc: "Read a stack of packing lists at once \u2014 Excel straight off the sheet, PDF through the same scanner the single-file screen uses. The result downloads as one spreadsheet, which Packing List Import reads back in a single upload. Check the last column before you use it: it flags any lot whose case list does not match the package count it declares.",
    plrDupCase: (code, files) => `case ${code} appears twice under this same reference (${files})`,
    plrRefClientClash: (list) => `this reference is given two different clients: ${list}`,
    plrRefSiteClash: (list) => `this reference is given two different sites: ${list}`,
    plrChooseBtn: "Choose packing lists\u2026",
    plrReading: (name) => `Reading ${name}\u2026`,
    plrSaveBtn: "Save table",
    plrSavedNote: (n) => `${n} row${n === 1 ? "" : "s"} saved`,
    plrSaveFailed: (err) => `could not be saved: ${err}`,
    plrRestored: (n, at) => `${n} row${n === 1 ? "" : "s"} restored from ${at || "an earlier session"}`,
    plrClearConfirm: (n) => `Clear all ${n} row${n === 1 ? "" : "s"}?\n\nAnything you have edited here and not downloaded will be lost.`,
    plrExportBtn: "Download spreadsheet",
    plrClearBtn: "Start over",
    plrTooBig: (kb) => `too big to scan at ${kb} KB \u2014 split it, or use the Excel version if there is one`,
    plrCount: (rows, files, off) => `${rows} lot${rows === 1 ? "" : "s"} from ${files} file${files === 1 ? "" : "s"}${off ? ` \u00b7 ${off} where the packages and the cases disagree` : ""}`,
    navLedger: "Storage ledger",
    ledgerTitle: "Storage ledger",
    ledgerDesc: "Every arrival and delivery the depot has recorded, a site at a time, in date order with a running balance \u2014 the same shape as a client's own store list, so the two can be read side by side. Charges and free days are not shown here; those follow the billing rules rather than the movements.",
    legacyReturnedNote: (job) => `cases returned on job ${job}`,
    ledgerNoDate: "NO DATE",
    ledgerAddDateHint: "Set the date this movement happened \u2014 it will then count in the balance",
    ledgerNotCounted: "not counted",
    ledgerUndatedTotal: (n) => `${n} movement${n === 1 ? "" : "s"} with no date \u2014 shown in red on the sites below, and left out of every balance until dated.`,
    ledgerAsOfLabel: "As of",
    ledgerAsOfToday: "Today",
    ledgerTotalOnHand: (d) => `packages on hand${d ? ` as at ${d}` : ""}, across every site`,
    ledgerSetAside: (later, undated) => [
      later ? `${later} movement${later === 1 ? "" : "s"} dated after this day` : "",
      undated ? `${undated} with no date` : "",
    ].filter(Boolean).join(" \u00b7 ") + " set aside, and not counted in the balance.",
    ledgerExportBtn: "Export ledger",
    ledgerSiteLine: (i, o, b) => `${i} in \u00b7 ${o} out \u00b7 ${b} on hand`,
    ledgerColDate: "Date",
    ledgerColJob: "FC job no.",
    ledgerColDir: "In / out",
    ledgerColType: "Type",
    ledgerColIn: "Pkgs in",
    ledgerColOut: "Pkgs out",
    ledgerColBalance: "Balance",
    storeListTitle: "Against the client's storage ledger",
    storeListDesc: "Load a client's inventory store list \u2014 the one with a sheet per site and a row for every movement in and out. Each job number in it is set against what the depot recorded under the same number, and only the rows that disagree are shown. A ledger that stops before this month will show this month's jobs as missing; that is expected, not an error.",
    storeListChooseBtn: "Choose store list\u2026",
    storeListLoaded: (name, n, sites) => `${name} \u2014 ${n} movements across ${sites} site${sites === 1 ? "" : "s"}`,
    storeListUnreadable: (err) => `That file could not be read: ${err}`,
    storeListSiteLine: (i, o, l, ai, ao, al) => `ledger ${i} in, ${o} out, ${l} left \u00b7 depot ${ai} in, ${ao} out, ${al} left`,
    storeListSiteAgrees: (n) => `All ${n} movements agree with the depot's records.`,
    storeListColJob: "FC job no.",
    storeListColDir: "In / out",
    storeListColDm: "DM on the ledger",
    storeListColSheet: "Ledger",
    storeListColApp: "Depot",
    storeListColDiff: "Difference",
    checkInsTitle: "Check-ins against the paperwork",
    checkInsDesc: "Every check-in the depot has taken, grouped by the reference it came in under. A reference states its own package count, so where the entries under it add up to more, the same consignment has been checked in twice \u2014 usually once as a whole lot and once split per lift.",
    checkInsOverBanner: (n) => `${n} package${n === 1 ? " is" : "s are"} counted more than once. Reversing a check-in puts its cases back to Incoming, where they can be checked in again properly.`,
    checkInsHeldVsStated: (held, stated, left) => `${held} checked in against ${stated} on the paperwork \u00b7 ${left} still at the depot`,
    checkInsHeldOnly: (held, left) => `${held} checked in \u00b7 ${left} still at the depot \u00b7 the reference states no package count, so nothing to check against`,
    checkInsOverBy: (n) => `${n} too many`,
    checkInsRowCount: (n) => `${n} check-in${n === 1 ? "" : "s"}`,
    checkInsSearchLabel: "Find a check-in",
    checkInsSearchPlaceholder: "Entry, shipment, lift or reference \u2014 e.g. INC-0559",
    checkInsColEntry: "Entry",
    checkInsColCode: "Lift / order",
    checkInsColArrived: "Arrived",
    checkInsColUnits: "Cases",
    checkInsColLeft: "Left",
    checkInsColSource: "Came in from",
    checkInsReverseBtn: "Reverse this check-in",
    checkInsMergeLabel: "Merge into\u2026",
    checkInsMergedNote: (from) => `moved from ${from}`,
    checkInsMergeConfirm: (from, keep, n, dels) => `Merge ${from} into ${keep}?\n\n${dels ? `Its ${dels} deliver${dels === 1 ? "y is" : "ies are"} moved onto ${keep} first, so nothing is lost. ` : ""}${from}'s check-in is then reversed and its ${n} case${n === 1 ? "" : "s"} go back to Incoming.\n\nThis cannot be undone from here.`,
    checkInsReverseHasDeliveries: (id, n) => `Careful \u2014 ${id} has ${n} deliver${n === 1 ? "y" : "ies"} recorded against it.\n\nUndoing its only check-in leaves those deliveries with no stock behind them. If this entry is the duplicate, note which deliveries it holds first, so they can be re-recorded against the entry you are keeping.\n\nCarry on anyway?`,
    checkInsReverseConfirm: (id, n, gone) => `Reverse this check-in?\n\n${n} case${n === 1 ? "" : "s"} go back to Incoming, waiting to be checked in again.${gone ? `\n\n${id} holds nothing else, so the entry is removed with it.` : ""}\n\nThis cannot be undone from here.`,

    depotOverviewTitle: "Depot Overview",
    depotOverviewItemsLabel: "item(s)",

    newEntryManual: "Manual",
    newEntryImport: "Import",

    moveCasesLabel: "Move Cases to Another Entry",
    moveCasesHint: "For cases filed under the wrong lift. Pick them, choose where they belong, and they move across with their arrival batch. Applies straight away, to both entries. Cases already delivered aren't shown - they belong to a delivery record.",
    moveCasesDestPlaceholder: "Move to which entry...",
    moveCasesBtn: "Move Cases",
    moveCasesDoneMsg: (codes, dest) => `Moved ${codes} to ${dest}.`,
    splitArrivalLabel: "Split Arrival — Cases Arriving on Different Days",
    splitArrivalHint: "If this packing list isn't all devanned / delivered to the depot on the same day, record each batch here with its own date and type. The main Depot Arrival Date will follow the earliest batch automatically.",
    splitArrivalCasesCol: "Cases",
    splitArrivalAddBtn: "Add Arrival Batch",
    splitArrivalSelectHint: "Tap the cases that arrived in this batch, then add it.",
    splitArrivalAllAssignedMsg: "All cases are assigned to an arrival batch.",
    badgePartialArrival: (a, b) => `AT DEPOT ${a}/${b} · MORE ARRIVING`,
    notYetArrivedTag: "not yet arrived",
    notYetArrivedHint: "This case has not yet arrived at the depot, so it can't be delivered out yet.",
    pendingArrivalNotice: (n) => `${n} case(s) have not yet arrived at the depot and can't be selected for delivery. Record their arrival batch on the Edit screen when they land.`,

    tabFreeStorage: "Free Storage",
    freeStorageTitle: "Free Storage Days",
    freeStorageDesc: "Standard free storage is 14 days from each devan / CFS arrival (Schindler: 21 days). For projects with a special arrangement — e.g. Otis, which varies by project — add a rule below. The rule applies when the project or construction site name contains the text you enter, and overrides the client default.",
    fFreeProject: "Project name contains",
    fFreeProjectHint: "e.g. Kwu Tung North Area 19",
    fFreeDays: "Free days",
    freeStorageAddBtn: "Add Rule",
    freeColProject: "Project match",
    freeColDays: "Free days",
    freeStorageNoneMsg: "No project rules yet — standard 14 days applies (Schindler: 21 days).",
    tabPricing: "CBM Pricing",
    pricingTitle: "CBM Rate per Client",
    pricingDesc: "Storage in the Warehouse and Open Yard is billed by CBM. These rates apply per CBM of storage; KG is only used when calculating Delivery and CFS charges. Leave a field blank to use the default rate shown as its placeholder.",
    pricingColClient: "Client",
    pricingColRate: "Rate",
    pricingPerCbm: "/ CBM",
    pricingResetBtn: "Reset to default",
    tabLegacy: "Legacy Uploads",
    legacyUploadTitle: "Legacy Upload",
    legacyUploadDesc: "Bulk-upload old Devan, CFS, Delivery, Shifting, Hoisting and other job sheet files. Best done in two rounds: upload all Devan/CFS files first (these create the storage arrivals), then upload Delivery files afterward \u2014 each one automatically finds and closes the arrival it refers to. Every file type is archived and kept searchable below regardless.",
    legacyChooseFilesBtn: "Choose Files\u2026",
    legacyDocType: "Doc Type",
    legacyProjectSite: "Project / Site",
    legacyUnitCode: "Unit / Lift No.",
    legacyPkgs: "Packages",
    legacyWeightKg: "Weight (kg)",
    legacyCbm: "CBM",
    legacyAlreadyDelivered: "Already delivered \u2014 close this job immediately instead of leaving it at the depot",
    legacyProcessBtn: (n) => `Process ${n} File${n === 1 ? "" : "s"}`,
    legacyProcessingMsg: "Processing\u2026",
    legacyResultsMsg: (archived, created, delivered, enriched, checkedIn) => `Archived ${archived} file${archived === 1 ? "" : "s"}${created > 0 ? ` \u2014 created ${created} inventory ${created === 1 ? "entry" : "entries"}` : ""}${checkedIn > 0 ? ` \u2014 checked in ${checkedIn} case group${checkedIn === 1 ? "" : "s"} against Incoming` : ""}${enriched > 0 ? ` \u2014 matched ${enriched} to existing ${enriched === 1 ? "entry" : "entries"} and added its details` : ""}${delivered > 0 ? ` \u2014 recorded ${delivered} deliver${delivered === 1 ? "y" : "ies"} against existing arrivals` : ""}.`,
    legacyImportedNote: (name) => `Imported from legacy file: ${name}`,
    legacyAutoClosedNote: "Auto-closed on legacy import (marked already delivered).",
    legacyBacklogTitle: "Backlog",
    legacyBacklogDesc: "Every file uploaded through Legacy Upload, whether it created an inventory entry or was archived only.",
    legacyBacklogNoneMsg: "No legacy files uploaded yet.",
    legacyBacklogNoMatchMsg: "No uploaded files match these filters.",
    legacyBacklogSortLabel: "Order",
    legacyBacklogSortRecent: "Most recent first",
    legacyBacklogSortJobNo: "By job number",
    legacyBacklogCount: (shown, total) => shown === total ? `${total} file${total === 1 ? "" : "s"}` : `${shown} of ${total} files`,
    legacyColFile: "File",
    legacyColLinked: "Linked Entry",
    legacyArchivedOnly: "Archived only",
    legacyEditLinkedHint: "The fields above fix this archive listing. The depot records this file created are below \u2014 correcting them here changes the linked FS-#### entries directly.",
    legacyLinkedRecordsLabel: "Records this file created",
    legacyLinkedRecordsHint: "A weight or volume typed here is taken as stated for that lot, and replaces whatever the sheet was read as. Leave both blank to fall back to the packing-list figures. Changing the entry moves the whole record onto that lot, which is how a job sheet filed against the wrong lift gets put right. To change which cases are involved, open the entry in Inventory.",
    legacyRecordEntryLabel: "Entry",
    legacyDeleteTitle: "Delete this file from the archive. Reversing also undoes what it did to the depot:",
    legacyDeleteArrivalLine: (label, n, cases) => `${label}: ${n} arrival${n === 1 ? "" : "s"} covering ${cases} case${cases === 1 ? "" : "s"} lifted off.`,
    legacyDeleteDeliveryLine: (label, n, cases) => `${label}: ${n} deliver${n === 1 ? "y" : "ies"} covering ${cases} case${cases === 1 ? "" : "s"} lifted off \u2014 those cases go back into store.`,
    legacyDeleteEntryRemoved: (label) => `${label} was created by this file and holds nothing else, so it is removed.`,
    legacyDeleteEntryKept: (arrivals, deliveries) => `The entry stays \u2014 ${arrivals} other arrival${arrivals === 1 ? "" : "s"} and ${deliveries} deliver${deliveries === 1 ? "y" : "ies"} remain on it.`,
    legacyDeleteIncomingLine: (incId, n) => `Incoming ${incId}: ${n} case${n === 1 ? "" : "s"} go back to waiting for check-in.`,
    legacyDeleteBlockedMsg: (label, n) => `${label} was created by this file but has since been delivered off (${n} deliver${n === 1 ? "y" : "ies"}), so it is left in place. Reverse those deliveries first if you want it gone.`,
    legacyDeleteStrandedMsg: (labels) => `Nothing of this file's making is still on ${labels}, so those are left alone.`,
    legacyDeleteNothingMsg: "This file created no depot records, so there is nothing to reverse \u2014 only the archive listing will go.",
    legacyDeleteHint: "Site names, SS/D.O. numbers and other details this file filled in on existing entries are not undone. Cases returned to Incoming can be checked in again from the Incoming tab.",
    invSelectedCount: (n, pkgs) => `${n} selected \u00b7 ${pkgs} package${pkgs === 1 ? "" : "s"}`,
    invBulkDeleteBtn: "Delete selected",
    invBulkDeleteConfirm: (n, pkgs, delivered) => `Permanently delete ${n} inventory ${n === 1 ? "entry" : "entries"} and their ${pkgs} package${pkgs === 1 ? "" : "s"}?\n\n${delivered ? `${delivered} of them ${delivered === 1 ? "has" : "have"} deliveries recorded against them, and those go too.\n\n` : ""}Export first if you have not. This cannot be undone.`,
    incomingSelectAll: (n) => `Select all ${n} shown`,
    incomingSelectedCount: (n, cases) => `${n} selected \u00b7 ${cases} case${cases === 1 ? "" : "s"}`,
    incomingBulkDeleteBtn: "Delete selected",
    incomingBulkDeleteConfirm: (n, cases, linked) => `Delete ${n} shipment${n === 1 ? "" : "s"} and their ${cases} case${cases === 1 ? "" : "s"}?\n\n${linked ? `${linked} of them ${linked === 1 ? "has" : "have"} already been checked into an inventory entry \u2014 those entries stay, but will no longer have a packing list behind them.\n\n` : ""}This cannot be undone.`,
    legacySelectAllHint: "Select every file the filters above are showing",
    legacySelectedCount: (n) => `${n} file${n === 1 ? "" : "s"} selected`,
    legacyBulkReverseBtn: "Reverse records and delete",
    legacyBulkDeleteBtn: "Delete listings only",
    legacyClearSelection: "Clear selection",
    legacyBulkDeleteConfirm: (n) => `Delete ${n} archive listing${n === 1 ? "" : "s"}?\n\nThe depot records they created stay exactly as they are \u2014 only the archived files go.\n\nThis cannot be undone.`,
    legacyBulkReverseConfirm: (n, entries, removed) => `Reverse and delete ${n} file${n === 1 ? "" : "s"}?\n\n${entries} inventory ${entries === 1 ? "entry is" : "entries are"} affected, of which ${removed} will be removed outright. Cases return to Incoming.\n\nExport your inventory before doing this. It cannot be undone.`,
    legacyDeleteReverseBtn: "Reverse records and delete file",
    legacyDeleteKeepRecordsBtn: "Delete file, keep records",
    legacyDeleteListingOnlyBtn: "Delete file",
    legacyRecordMoveNote: (from, to, cases, kept) => `Moves off ${from} onto ${to} on save.${cases ? (kept === cases ? ` All ${cases} case${cases === 1 ? "" : "s"} carry over.` : ` ${kept} of ${cases} case numbers exist on ${to}; the other ${cases - kept} will be dropped, so check the weight and volume above.`) : ""}`,
    legacyClientUnresolved: "\u2014 select client \u2014",
    legacyClientRequiredSummaryMsg: "One or more files don't have a recognized client \u2014 select the correct client for each file before processing.",
    legacyDeliveredFrom: (id) => `Delivered from ${id}`,
    legacyUnmatchedReferral: (job) => `No arrival found for job ${job}`,
    legacyUnmatchedHint: "This delivery refers to a job number that isn't in the system yet - upload its Devan/CFS file first, or link it manually later.",
    legacyProjectSiteEn: "Project / Site (English)",
    legacyProjectSiteZh: "Project / Site (Chinese)",
    legacySiteRequiredMsg: "Enter at least one of English or Chinese site name.",
    legacySiteRequiredSummaryMsg: "One or more files are missing a site name in both languages \u2014 fill in at least one (English or Chinese) for each file before processing.",
    legacyScanningMsg: "Reading files\u2026",
    legacyAutoDetectHint: "Excel files (.xlsx/.xls/.csv) are read directly, and scanned PDFs are read by scanning the page \u2014 client, site, job number, date, SS/D.O. info, referral blocks, case numbers and totals all come through. Review the pre-filled fields below before processing; a scan is not as certain as a spreadsheet. Images still need manual entry.",
    legacyAutoDetectedTag: "Auto-detected from file \u2014 please check",
    legacyReferLine: (job, date) => `refers to job no. ${job} on ${date}`,
    legacyReferJobNoLabel: "Refers to Arrival Job No.",
    legacyEnrichedNote: (name) => `Enriched from legacy file: ${name}`,
    legacyMatchedIncoming: (id) => `Matched to Incoming ${id} \u2014 select which of its cases arrived in this file`,
    legacyNoPackingListTitle: "No packing list on file for this site",
    legacyPackingListLookedFor: (client, site) => `Looked for a shipment waiting under ${client || "(no client)"} at ${site || "(no site)"} and found none.`,
    legacyPackingListNearMiss: (n, list) => `${n} shipment${n === 1 ? " is" : "s are"} waiting to be checked in under other names \u2014 ${list}. If one of those is this shipment, correct the client or site above so it matches, rather than building a second packing list here.`,
    legacyNoPackingListDesc: "Nothing here to check these cases into. This sheet names its own lots and cases, so a packing list can be built straight from it \u2014 the lot's stated weight and volume are spread evenly across its cases, which is all the sheet claims. Every case stays editable in Incoming afterwards.",
    legacyCreatePackingListBtn: (n) => `Create packing list \u2014 ${n} lot${n === 1 ? "" : "s"}`,
    legacyPackingListFromSheetNote: (src) => `Packing list entered from job sheet ${src}`,
    legacyMatchedIncomingCount: (n) => `Matched ${n} Incoming shipment${n === 1 ? "" : "s"} at this site \u2014 select which cases from each arrived in this file`,
    legacyReferJobNoHint: "Optional \u2014 narrows the match to one specific arrival job number. Leave blank to match by client + site instead (can find several).",
    legacyOversizeLabel: "Oversize",
    legacyOversizeCbmPlaceholder: "CBM",
    legacyOversizeHint: "Bills at the oversize rate for this client, using this CBM as the basis.",
    legacyMatchedItemsCount: (n) => `Matched ${n} inventory ${n === 1 ? "entry" : "entries"} with cases still at the depot \u2014 select which left in this delivery`,
    legacyMatchedItemsReturn: (n) => `Matched ${n} inventory ${n === 1 ? "entry" : "entries"} with cases out at site \u2014 select which came back on this return`,
    legacyMatchedItemReturn: (id) => `Returning to ${id}`,
    legacyTypeSelectPlaceholder: "e.g. 1,3-5,7",
    legacyTypeSelectBtn: "Add",
    legacySelectedTotals: (count, kg, cbm) => `Selected: ${count} pkg${count === 1 ? "" : "s"} \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacySelectedTotalsGrand: (count, kg, cbm) => `Total selected across all entries: ${count} pkg${count === 1 ? "" : "s"} \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacyDeliveryDeclaredLabel: "On this delivery sheet:",
    legacyOversizeCasesPh: "e.g. 14/21",
    legacyOversizeCaseCol: "Case",
    legacyOversizeAdd: "+ Add oversize case",
    legacyOversizeRemove: "Remove",
    legacyOversizeTotal: (n, cbm) => `${n} oversize cases \u00b7 ${cbm} cbm total \u00b7 each bills at its own tier`,
    legacyOversizeDeliveryHint: "Prints in the OVERSIZE CASES box on this delivery sheet",
    legacyDeliveryDeclaredGap: (declaredKg, listedKg, pct, heavier) =>
      `Sheet declares ${declaredKg} kg; the cases selected come to ${listedKg} kg on the packing list (${heavier ? "+" : "\u2212"}${pct}%). The sheet's figure is recorded.`,
    legacySheetTotalLabel: (lots, ref) => `Total stated on this sheet${ref ? ` \u00b7 ${ref}` : ""} \u2014 covers ${lots}`,
    legacySheetTotalJob: (job) => `Job ${job}`,
    legacySheetTotalHint: "The sheet gives one total for these lots. Edit it here; each lot's share below is worked out from it, pro-rata on the packing list.",
    legacyDeclaredShareNote: (kg, cbm) => `Share of the sheet total: ${kg} kg \u00b7 ${cbm} cbm`,
    incomingDeclaredLabel: "Totals on the Devan/CFS sheet (optional)",
    legacyDeclaredLabel: "On this sheet:",
    legacyDeclaredHint: "Sheet totals are used for storage, handling and billing. The packing-list weight stays on record case by case.",
    legacyDeclaredKgGap: (kg, heavier, pct) => heavier
      ? `Sheet is ${kg} kg heavier than the packing list (+${pct}%) \u2014 expected, it includes crating and skids on truss sections.`
      : `Sheet is ${kg} kg lighter than the packing list (\u2212${pct}%) \u2014 worth checking before you process it.`,
    legacyDeclaredPkgsGap: (declared, selected) => `Sheet counts ${declared} package${declared === 1 ? "" : "s"}, you selected ${selected}.`,
    fWeightFromSheet: "From the Devan/CFS sheet",
    fWeightSheetShare: "Estimated \u2014 share of a sheet total covering several lots",
    fWeightPackingList: (kg) => `Packing list: ${kg} kg`,
    fVolumePackingList: (cbm) => `Packing list: ${cbm} cbm`,
    fVolumeFromOversize: "Oversize CBM \u2014 sets the billing tier, so it outranks the sheet total",
    legacyMatchedItem: (id) => `Delivering from ${id}`,
    legacyArrivalStaysOpenHint: "Stays open at the depot until a matching Delivery file is uploaded (or you record a delivery for it normally).",
    legacyNoReferralHint: "No \"Ref Job no.\" line detected \u2014 enter the arrival's job number manually, or this file will only be archived.",
    legacySheetCasesNote: (mark, n) => `Sheet marks ${mark} \u2014 ${n} case${n === 1 ? "" : "s"} pre-selected`,
    legacyCasesFormatClash: (sheet, list) => `Not one of them matched, so nothing could be pre-selected. The sheet asks for ${sheet}\u2026 while the packing list holds ${list}\u2026 \u2014 the same cases written differently. Re-import the packing list to rebuild its markings, or tick the cases by hand this once.`,
    legacySheetCasesMissing: (list) => `case ${list} not at the depot`,
    legacySheetCasesElsewhere: (list, more) => `but checked in elsewhere: ${list}${more ? `, and ${more} more` : ""} \u2014 deliver from that entry instead, or move the cases across in Inventory`,
    legacyReplaceCasesHint: (n) => `Same number of cases, but ${n} numbered differently here than on this sheet. Either source can be wrong \u2014 a scan can drop a digit, a typed sheet can carry a typo \u2014 so check the paper before changing anything:`,
    legacyReplaceCasesBtn: "Renumber this shipment to match the sheet",
    legacyScannedFromPdf: "Read by scanning the page \u2014 check the fields and case numbers against the paper before processing.",
    legacyScanFailed: (why) => `Could not read this PDF (${why}). Fill the fields in by hand, or upload the Excel original if there is one.`,
    legacyCasesArePicture: "This sheet's case list is a picture pasted into it, not typed cells, so nothing can be read out of it. Type the markings into the case box below, or into the Case Numbers column of the import spreadsheet.",
    legacyCaseCountMismatch: (lot, stated, listed) => `${lot}: the sheet says ${stated} PKGS but lists ${listed} case number${listed === 1 ? "" : "s"} \u2014 check the C/S No. list before processing.`,
    legacyIncomingCasesMissing: (list) => `case ${list} is not on this shipment`,
    legacyCasesFoundInMore: (n) => `\u2026 and ${n} more shipment${n === 1 ? "" : "s"} on this lot.`,
    legacySheetCasesAmbiguous: (mark, ids) => `Sheet marks ${mark}, but ${ids} both answer to this lot \u2014 nothing pre-selected, pick the cases yourself.`,
    legacyCasesFoundIn: (codes, id, unit) => `\u2192 ${id} \u00b7 ${unit} carries ${codes} under the same order number. Case numbers alone don't prove it's the same box \u2014 confirm against the paperwork before checking in there.`,
    legacyMisfiledCases: (codes, from) => `Case ${codes} sits under ${from}, but the lot size on the case number says it belongs here \u2014 the packing list's lift column had it wrong.`,
    legacyMisfiledFixBtn: "Move it here",
    legacySomeArrivalsUnmatched: (jobs) => `No inventory entry found for arrival job ${jobs} \u2014 that part of this sheet won't be delivered. Upload its Devan/CFS file first, or correct the number above.`,
    legacyArrivalLaterInBatch: (job, file) => `The arrival for job ${job} is in this batch (${file}) and has not been processed yet. Arrivals are created before deliveries when you press Process, so this delivery will find it \u2014 nothing to do.`,
    legacyNoArrivalFoundHint: (job) => `No inventory entry found with job number ${job} \u2014 check the number is correct, or upload that Devan/CFS file first. This file will still be archived, but no delivery will be recorded against it.`,
  },
  zh: {
    appSubtitle: "倉庫及貨物存倉表",
    navDashboard: "總覽",
    navInventory: "存倉記錄",
    navNewEntry: "新增記錄",
    navDeliveries: "送貨記錄",
    navDuplicates: "重複記錄",
    navDuplicatesCount: (n) => `重複記錄（${n}）`,
    navImport: "匯入",
    themeToggleLabel: "切換深色模式",
    langToggleLabel: "EN",
    loadingMsg: "正在載入倉存記錄…",
    saveErrorMsg: "儲存失敗，請重試。",
    conflictTitle: "已有其他使用者先行儲存",
    conflictBody: "此頁開啟期間，另一位使用者已儲存變更，因此剛才的修改並未寫入 \u2014 否則會覆蓋對方的紀錄。下方顯示的是目前實際已儲存的內容，請重新輸入。",
    conflictDismiss: "知道了",

    statAtDepot: "在倉",
    statPending: "待提取",
    statBillable: "現正收費",
    statLfd: "截關期提示",
    dupBanner: (n) => `發現 ${n} 組可能重複的記錄 — 客戶、項目、單位編號及抵倉日期相同。`,
    reviewDuplicatesBtn: "查看重複記錄",
    lfdSectionTitle: "貨櫃碼頭截關期 — 需要處理",
    billableSectionTitle: "現正收費之存倉項目",
    billableEmptyMsg: "暫時沒有超過14天免費存倉期的項目。",
    sinceLabel: "由",

    searchLabel: "搜尋",
    searchPlaceholder: "項目、客戶、編號、SHK編號、件號、工單號",
    clientLabel: "客戶",
    statusLabel: "狀態",
    statusAll: "全部",
    statusPending: "待提取",
    statusAtDepot: "在倉",
    statusPartial: "部分已送貨",
    statusDelivered: "已送貨",
    depotLabel: "貨倉",
    exportBtn: (n) => `匯出Excel（${n}）`,
    newEntryBtn: "+ 新增記錄",
    colId: "編號",
    colClient: "客戶",
    colProjectSite: "項目/地盤",
    colUnit: "單位編號",
    deliverySearchPlaceholder: "客戶、項目、單位編號、工單號或件號",
    colDepot: "貨倉",
    colDepotArrival: "抵倉日期",
    colStatus: "狀態",
    noRecordsMsg: "沒有符合的記錄。請新增記錄開始使用。",
    editBtn: "編輯",
    deleteBtn: "刪除",
    deliverBtn: "送貨",
    duplicateBadge: "重複？",

    titleNew: "新增倉存記錄",
    titleEdit: "編輯倉存記錄",
    fClient: "客戶",
    fProject: "項目",
    fProjectEn: "項目／地盤（英文）",
    fProjectZh: "項目／地盤（中文）",
    fProjectPlaceholder: "例如：港鐵油塘站現代化工程",
    fInvoiceNo: "發票編號",
    fInvoiceHint: "開發票後填寫",
    sectionSite: "地盤資料",
    fSiteName: "地盤名稱",
    fSiteHint: "英文或繁體中文皆可，毋須兩者兼填",
    fSitePlaceholder: "例如：油塘站 or MTR Yau Tong Station",
    sectionCargo: "貨物詳情",
    fItemType: "貨物類型",
    fPackages: "件數",
    fUnitCode: "扶手電梯/升降機編號",
    fUnitCodeHint: "例如：L1、E02",
    fDescription: "描述",
    fDescriptionHint: "例如：單位編號／設備描述",
    fReference: "參考編號／SHK編號",
    fReferenceHint: "適用於迅達（Schindler）之參考編號",
    fWeight: "重量（公斤）",
    fVolume: "體積（立方米）",
    f20: "20呎貨櫃數量",
    fZone: "地區",
    fZoneHint: "影響已有收費表客戶的拆櫃/送貨收費",
    zoneUrban: "市區",
    zoneLantau: "大嶼山／東涌",
    fCargoType: "貨物類型",
    fCargoTypeHint: "留空（自動）將按描述自動判斷",
    cargoTypeAuto: "自動判斷",
    cargoTypeElevator: "升降機",
    cargoTypeEscalator: "電扶梯",
    f40: "40呎貨櫃數量",
    sectionArrival: "抵達及倉存資料",
    fArrivingType: "抵達方式",
    fArrivingTypeHint: "Devan（拆櫃）= 由我方自行拆櫃；CFS = 客戶直接送來散件，毋須拆櫃。",
    fDepot: "貨倉",
    fDepotHint: "貨物存放之貨倉",
    fDepotLocation: "倉位/存放位置",
    fTerminalArrival: "貨櫃碼頭抵港日期",
    fTerminalArrivalHint: "預計抵達貨櫃碼頭日期",
    fTerminalLFD: "貨櫃碼頭截關期",
    fTerminalLFDHint: "須於此日期前提取貨物，否則產生額外費用",
    fConfirmedCollection: "確認提貨日期",
    fConfirmedCollectionHint: "我方前往貨櫃碼頭提貨之日期",
    fDepotArrival: "抵倉日期",
    fDepotArrivalHint: "開始計算14天免費存倉期",
    fPlannedDelivery: "預計送貨日期",
    fPlannedDeliveryHint: "僅供預算之用 — 實際送貨請於「送貨記錄」分頁記錄",
    deliveryProgress: (del, tot, count) => `已送出 ${tot} 件中之 ${del} 件，共 ${count} 次送貨記錄，最近一次為 `,
    deliveryProgressManage: "請於「送貨記錄」分頁管理送貨情況。",
    fNotes: "備註",
    saveBtn: "儲存記錄",
    cancelBtn: "取消",

    deliveryTitlePrefix: "記錄送貨 —",
    combinedDeliveryTitle: (n) => `記錄合併送貨 — 共${n}項記錄`,
    combinedHistoryHiddenNote: "此處不會顯示這些記錄各自的送貨歷史 — 如需查看或取消個別記錄的送貨，請從存倉記錄個別打開。",
    recordCombinedBtn: (n) => n > 0 ? `為已選項目記錄送貨 (${n})` : "為已選項目記錄送貨",
    combinedSelectionHint: "勾選所有要一同送出的記錄（例如兩個不同批次但同日送往同一地盤），然後合併記錄為一張工單。",
    selectedTag: "已包含 ✓",
    includeInDeliveryBtn: "加入此次送貨",
    addCombinedDeliveryBtn: "記錄合併送貨",
    combinedPrintLabel: "合併送貨工單",
    addMoreBatchesBtn: "+ 加入更多批次到此送貨記錄",
    addMoreBatchesSearchPlaceholder: "按單位編號或工單號搜尋",
    addMoreBatchesScopeNote: (site) => `只顯示同屬 ${site} 的其他記錄。`,
    addMoreBatchesNoneMsg: "沒有其他屬於同一地盤的記錄可供加入。",
    plannedWasText: (date) => ` · 預計送貨日期為 ${date}`,
    progressOf: "／",
    progressDeliveredSoFar: "件已送出",
    progressRemaining: "件尚在倉",
    colDate: "日期",
    colQty: "數量",
    colDeliveredTo: "送達地點",
    colReceivedBy: "接收人",
    removeBtn: "移除",
    fDeliveryDate: "送貨日期",
    fDeliveryDateHint: "全部送出後將以此日期結算存倉期",
    fQty: "送貨數量",
    fQtyHint: (r) => `尚餘 ${r} 件`,
    fDeliveredTo: "送達地點",
    fDeliveredToHint: "預設為地盤名稱",
    fReceivedBy: "接收人",
    fReceivedByHint: "留空則使用此項目之落單人",
    overshootMsg: (r) => `此記錄現時只剩 ${r} 件在倉，請減少數量。`,
    addDeliveryBtn: "新增送貨記錄",
    closeBtn: "關閉",
    allDeliveredMsg: "此記錄之貨物已全部送出。",
    selectItemMsg: "請選擇項目以記錄送貨。多於一件之貨物（例如一批電梯零件）可分批送貨 — 每次記錄將扣減在倉數量。",
    nothingAtDepotMsg: "目前貨倉內沒有貨物。",
    recordDeliveryBtn: "記錄送貨",

    noneFoundMsg: "未發現重複記錄。當客戶、項目、單位編號、貨物類型、抵倉日期及發票編號均與另一記錄相同時，系統會標示為重複。",
    matchingEntries: (n) => `${n} 個相符記錄`,
    dupSharedLine: (arrival, type, jobs) => `同一抵倉日 ${arrival} \u00b7 ${type} \u00b7 工單號 ${jobs}`,
    dupColCases: "貨箱",
    dupColDeliveries: "送貨",
    dupFlatCount: (n) => `${n}（無貨箱清單）`,
    dupSourceLabel: "來源",
    dupHasDeliveriesWarn: (n) => `此記錄有 ${n} 筆送貨記錄 \u2014 刪除將一併移除`,
    dupDiffHint: "以顏色標示者為各記錄之間有差異的欄位；其餘欄位已相同，故被判定為重複。",
    deleteAllBtn: (n) => `全部刪除（${n}項）`,
    colInvoiceNo: "發票編號",
    colAddedOn: "新增日期",
    keepDeleteBtn: "保留此項，刪除其他",

    tabExcel: "上載Excel",
    tabPdf: "掃描PDF",
    pdfCaseCountMismatch: (lot, stated, read) => `${lot}：文件註明 ${stated} 件，但讀取到 ${read} 個件號，相差 ${Math.abs(stated - read)} 件 \u2014 請與紙本核對 C/S NO. 清單。系統只採用已讀取之 ${read} 件，不會自行補足差額。`,
    pdfWholeDocument: "此文件",
    pdfWeightMismatch: (stated, read) => `文件總重為 ${stated} 公斤，但讀取到之各件合計為 ${read} 公斤 \u2014 匯入前請與紙本核對下方批次。`,
    pdfCasesCorrected: (lot, changes) => `${lot}：已按文件嘜頭 (Shipping Marks) 修正件號 \u2014 ${changes}。重量及體積不變。`,
    pdfRepeatedLots: (lots) => `有多於一組名為 ${lots}。如文件將同一訂單分為不同批次送往不同電梯，請於匯入前分別命名，否則會合併為同一批次。`,
    pdfDocumentTotals: (cbm, kg) => `文件只提供整批總數 \u2014 ${cbm} CBM、${kg} 公斤 \u2014 並未按訂單分列，故系統不作分攤。如有資料請自行輸入各批次體積。`,
    pdfTerminalDatesFound: (eta, lastFree) => `此文件另有碼頭日期 \u2014 到港 ${eta}，免費倉期至 ${lastFree}。辦理到倉時請於記錄輸入；此並非裝箱單資料。`,
    tabManualPackingList: "手動輸入",
    manualLotBuilderLabel: "由工單建立批次",
    manualLotBuilderHint: "一張CFS工單通常包含多個批次。逐一填寫此列並按「新增批次」，各批次會列於下方並一併加入。按工單原文輸入 C/S NO.（1,2,3/3 \u00b7 1-12/23 \u00b7 1/2, 2/2），系統會自動建立各件並平均分配重量及體積。若件號並非純數字，可留空 C/S No.，改於下方逐件輸入。",
    manualOrderNoLabel: "訂單號",
    manualCaseSpecLabel: "貨箱編號",
    manualAddLotBtn: "新增批次",
    manualCaseSpecPreview: (n, codes) => `共 ${n} 件：${codes}`,
    manualPendingLotsLabel: (n) => `已備妥 ${n} 個批次`,
    manualUnnamedLot: "（未命名批次）",
    legacySameOrderWarn: (orders) => `訂單號 ${orders} 見於本地盤多於一批到貨記錄。同一訂單可分批付運，而每張裝箱單各自編號，故件號相同未必為同一箱 \u2014 亦可能是同一批貨重複記錄。請先核對文件再決定於哪一項辦理到倉。`,
    legacySameOrderSibling: (id, unit, n) => `${id} \u00b7 ${unit} \u00b7 尚有 ${n} 件未入倉`,
    legacySameOrderOnlyThere: (codes) => `該項有 ${codes}，此項則沒有`,
    manualPackingListOpenBtn: "+ 新增裝箱單",
    manualPackingListDesc: "適用於較舊、沒有裝箱單檔案或資料不完整的工作 \u2014 直接手動輸入件號清單。此操作會建立一項待到倉貨件，效果與上載檔案相同，可於拆櫃/CFS辦理到倉。",
    legacyManualEntryNote: "手動輸入 \u2014 此貨件沒有裝箱單檔案存檔。",
    excelTitle: "從Excel匯入",
    excelDesc: "系統會自動將欄位標題與倉存系統之欄位配對（例如「發票編號」、「抵倉日期」）。無法識別之欄位將被略過並於下方列出。",
    chooseFileBtn: "選擇檔案（.xlsx、.xls、.csv）",
    downloadTemplateBtn: "下載空白範本",
    selectedCount: (sel, tot) => `已選擇 ${sel} 項，共 ${tot} 項可匯入。`,
    selectAllBtn: "全選",
    clearBtn: "清除",
    legacyRangeTapHint: "點選一件，再按住 Shift 點選另一件，即可選取兩者之間全部。",
    selectNonDupBtn: "只選擇非重複項目",
    unmatchedMsg: "無法識別之欄位（已略過）：",
    prevColClient: "客戶",
    prevColProject: "項目",
    prevColItemType: "貨物類型",
    prevColDepot: "貨倉",
    prevColDepotArrival: "抵倉日期",
    prevColMatch: "配對結果",
    importBtn: (n) => `匯入 ${n} 項記錄`,
    discardBtn: "捨棄",
    pdfTitle: "掃描PDF裝箱單",
    pdfDesc: "上載客戶之裝箱單、送貨通知或Shipping List PDF — 任何格式皆可。系統會自動讀取升降機/批次分類、件號、重量及CBM，效果與Excel裝箱單匯入相同。",
    choosePdfBtn: "選擇PDF檔案",
    scanningMsg: "正在讀取文件…",
    pdfReadErrorMsg: "未能自動讀取此PDF。請檢查檔案，或於下方手動輸入資料。",
    pdfTruncatedMsg: "此文件內容非常密集（件數/項目眾多），自動讀取中途被截斷。請再試一次（有時第二次會成功），或於下方手動輸入資料。",
    pdfKeyWarning: "此功能使用您自己的Anthropic API金鑰，於下方輸入並只保存在此瀏覽器內。由於是直接從此頁面使用，任何開啟此檔案開發人員工具的人都有機會看到金鑰 — 適合內部信任員工測試，不適合對外派發。",
    pdfApiKeyLabel: "Anthropic API 金鑰",
    pdfApiKeyHint: "從 console.anthropic.com 取得 — 只保存在此瀏覽器",
    pdfSaveKeyBtn: "儲存金鑰",
    pdfKeySavedBadge: "金鑰已儲存",
    pdfNeedKeyMsg: "請於上方輸入並儲存API金鑰以啟用PDF掃描功能。",
    reviewWarningMsg: "請於儲存前核對以下自動讀取之資料 — 自動讀取或有遺漏或錯置。",
    dupWarningMsg: (match) => `此文件可能與現有記錄相符：${match.id}（${match.client} · ${match.project}）。`,
    excelErrorMsg: "未能讀取此檔案。請確保檔案為 .xlsx、.xls 或 .csv 格式。",
    excelNoRowsMsg: "此檔案內沒有資料列。",

    badgePendingCollection: "待提取",
    badgeLfdOverdue: (d) => `已過截關期 ${d} 日`,
    badgeLfdSoon: (d) => `截關期尚餘 ${d} 日`,
    badgeBillable: (d) => `計費中 · ${d}日`,
    badgeFree: (d) => `免費 · 尚餘${d}日`,
    badgePartial: (del, tot, extra) => `部分送貨 ${del}/${tot}${extra}`,
    badgeDelivered: (extra) => `已送貨${extra}`,
    badgeBilledSuffix: (d) => ` · 已收費${d}日`,
    badgeBillableSuffix: (d) => ` · ${d}日`,
    badgeDupOf: (id) => `重複於 ${id}`,
    badgeNew: "新項目",

    sectionPackages: "貨物件號清單",
    packagesHint: "選填 — 加入個別件號（例如 2A、13A），方便分開送貨。留空則只記錄總件數。",
    bulkAddLabel: "快速加入（以逗號或換行分隔）",
    bulkAddPlaceholder: "例如：1A, 2A, 3A, 12A, 13A",
    bulkAddBtn: "加入件號",
    packagesGenerateLabel: "共有幾件？",
    packagesGenerateBtn: "生成列數",
    packagesTotalLabel: (n) => `合計：${n} 件`,
    packageCodeCol: "件號",
    packageDescCol: "描述",
    packageWeightCol: "重量（公斤）",
    packageCbmCol: "CBM",
    removePackageBtn: "移除",
    packagesCountSummary: (n) => `已列出 ${n} 件`,
    noPackagesMsg: "尚未列出個別件號 — 將使用上方之總件數。",

    uploadModePackingList: "裝箱單（Excel／PDF）",
    uploadModeLegacy: "舊資料上載（CFS／拆櫃／送貨）",
    tabPackingList: "匯入客戶裝箱單",
    packingListTitle: "匯入客戶裝箱單",
    packingListDesc: "直接上載客戶原本的裝箱單（TK Elevator、Schindler、OTIS 等各自格式不同 — 系統會自動辨識欄位）。文件內每部升降機/每個批次將自動分成獨立記錄，並列出其貨物件號。",
    choosePackingListBtn: "選擇裝箱單檔案",
    packingListCaseCountWarn: (list) => `\u5df2\u8b80\u53d6\uff0c\u4f46\u8acb\u5148\u6838\u5c0d\uff1a${list}\u3002\u4ef6\u6578\u8207\u4ef6\u865f\u6578\u76ee\u4e0d\u7b26\uff0c\u5426\u5247\u5165\u5009\u4ef6\u6578\u5c07\u8207\u55ae\u64da\u4e0d\u4e00\u81f4\u3002`,
    packingListNoStructure: "未能在此檔案中識別裝箱單表格結構，可能是未見過的格式 — 請改用「上載Excel」，或手動新增記錄。",
    packingListDetectedTitle: (n) => `在此檔案中找到 ${n} 個升降機/批次`,
    packingListCommonFieldsTitle: "套用至以下全部項目",
    packingListApplyClient: "客戶",
    packingListApplyProject: "項目",
    packingListApplyDepot: "貨倉",
    packingListApplyDepotArrival: "抵倉日期（全部貨物一同抵達）",
    packingListApplyDepotLocation: "倉位/存放位置",
    packingListImportBtn: (n) => `匯入 ${n} 項記錄`,
    colLot: "升降機/批次",
    colPackages: "件數",
    colContainers: "貨櫃",
    colWeight: "重量（公斤）",
    selectByBatchLabel: "按到倉批次選取（此送貨可同時包含多於一個到倉批次）",
    selectCodesLabel: "選擇要送出的件號",
    noCodesRemainingMsg: "此記錄之個別件號已全部送出。",

    colCbm: "CBM",
    colKg: "KG",
    colJobNo: "快達單號",

    sectionJobSheet: "工單",
    fJobNumber: "快達單號",
    fJobNumberHint: "格式為年月+流水號，Devan/CFS/送貨共用同一組編號",
    generateJobNoBtn: "產生單號",
    fOrderedBy: "落單人",
    fPoNumber: "採購編號",
    fPoNumberHint: "由客戶提供",
    fJobRef: "地盤代號",
    fJobRefHint: "例如：KTN-002、GAGE-001",
    printJobSheetBtn: "列印工單／匯出PDF",
    printBtn: "列印／匯出PDF",
    closePreviewBtn: "關閉",

    jsTitle: "JOB SHEET",
    jsTitleZh: "工　單",
    jsFrom: "FROM",
    jsFromZh: "由",
    jsTo: "TO",
    jsToZh: "送",
    jsAccount: "ACCOUNT",
    jsAccountZh: "客戶",
    jsJobNo: "JOB NO.",
    jsJobNoZh: "快達單號",
    jsDate: "DATE",
    jsDateZh: "日期",
    jsOrderedBy: "ORDERED BY",
    jsOrderedByZh: "落單人",
    jsPoNo: "P.O. NO.",
    jsPoNoZh: "採購編號",
    jsJobRef: "JOB REF.",
    jsJobRefZh: "地盤代號",
    jsDescription: "DESCRIPTION",
    jsDescriptionZh: "貨物資料／工作程序",
    jsIssuedBy: "ISSUED BY",
    jsIssuedByZh: "出單人",
    jsTotal: "TOTAL",
    jsTotalZh: "共",
    jsPkgs: "PKGS",
    jsKgs: "KGS",
    jsCbm: "CBM",
    jsDevanFrom: (dep) => `DEVAN AT ${dep}`,
    jsCfsFrom: "CFS \u2014 客戶直接送到，毋須拆櫃",
    jsDeliveryType: "送貨 DELIVERY",
    jsDevanType: "拆櫃 DEVAN",
    jsCfsType: "CFS",
    jsSignatureLine: "客戶簽署確認 (按以上工作完成):",
    jsTemplateLabel: "工單類型",
    jsCfsFromPreset: "CFS 起運地點…",
    jsEditableHint: "虛線框可於列印前編輯。",
    jsOversizeLabel: "超長/超大件 Oversize Cases:",
    jsOversizePlaceholder: "每行 #件號 @X.XXCBM，最後一行 (合共 Total: X CBM)  xN 倍",
    jsOversizeNote: "超大件體積按上列倍數計算運費，僅適用於 Schindler 及 Chevalier 客戶。",
    deliveryHistoryLabel: "送貨記錄",
    signedDocArrivalLabel: "已簽收到倉工單 (拆櫃/CFS)",
    signedDocUploadBtn: "上載已簽工單",
    signedDocReplaceBtn: "更換已簽工單",
    signedDocViewBtn: "查看已簽工單",
    signedDocSavingMsg: "儲存中…",
    signedDocFailMsg: "未能儲存此檔案。相片會自動壓縮，但PDF須小於約3MB — 可改為上載已簽工單的相片。",
    fSsDoNo: "提單資料 SS/D.O. No.",
    fSsDoNoHint: "船名、航次及櫃號 — 會自動填入拆櫃/CFS工單的提單資料欄",
    jsEstimatedNote: "~ 重量／CBM為按比例估算，並非逐件過磅。",

    navDirectory: "目錄",
    tabSitesAccounts: "地盤及客戶",
    tabEmployees: "員工",

    dirTitle: "地盤及客戶目錄",
    dirDesc: "每個真實地盤只需登記一次 — 統一客戶、地盤代號及聯絡人資料，避免同一地盤因英文／中文名稱不同而被當作不同記錄。",
    dirAddBtn: "+ 新增地盤",
    fSiteEn: "地盤名稱（英文）",
    fSiteZh: "地盤名稱（中文）",
    fDirClient: "客戶",
    fDirJobRef: "地盤代號",
    fDirOrderedBy: "預設落單人",
    fDirOfficer: "負責客戶主任",
    dirColSite: "地盤",
    dirColClient: "客戶",
    dirColJobRef: "地盤代號",
    dirColOfficer: "客戶主任",
    dirColOrderedBy: "落單人",
    dirNoneMsg: "尚未新增任何地盤。",
    selectFromDirectory: "從目錄自動填寫",
    selectFromDirectoryPlaceholder: "— 選擇地盤以自動填寫 —",
    showOlderJobs: "顯示舊工程",
    showOlderJobsCount: (n) => `（已隱藏 ${n} 個已完成地盤）`,
    manualLinkedToDirectory: "已連結至目錄中的地盤 \u2014 中文名稱、地盤代號及落單人已自動填入。",
    saveNewSiteToDirectory: (name) => `將「${name}」儲存為目錄中的新地盤，日後匯入將自動識別`,

    siteTotalsTitle: "各地盤存倉之CBM及KG",
    siteRefsColRef: "參考編號（DM / SHK）",
    siteRefsColLots: "電梯",
    siteRefsColCases: "尚存倉之件號",
    siteTotalsColLastCfs: "最近CFS",
    siteTotalsColLastDevan: "最近拆櫃",
    siteTotalsColLastReturn: "最近退倉",
    siteTotalsColLastDelivery: "最近送貨",
    siteTotalsColSite: "地盤",
    siteTotalsColClient: "客戶",
    siteTotalsColPkgs: "餘下件數",
    siteTotalsColCbm: "剩餘CBM",
    siteTotalsColKg: "剩餘KG",
    siteTotalsNoneMsg: "目前貨倉內沒有貨物。",
    siteTotalsToggleHide: "隱藏",
    siteTotalsToggleShow: "顯示",

    empTitle: "員工",
    empDesc: "加入所有負責資料輸入、拆櫃／CFS或送貨的同事。職位可從清單選擇，亦可自行輸入新職位。",
    empAddBtn: "+ 新增員工",
    fEmpName: "姓名",
    fEmpRole: "職位",
    fEmpRolePlaceholder: "選擇或輸入職位",
    empColName: "姓名",
    empColRole: "職位",
    empNoneMsg: "尚未新增任何員工。",

    signedInAs: "登入身份",
    mobileMenuLabel: "選單",
    loginNameLabel: "姓名",
    loginNamePlaceholder: "請選擇你的姓名",
    loginPasswordLabel: "密碼",
    loginErrorMissing: "請輸入姓名及密碼。",
    loginErrorWrong: "姓名或密碼錯誤。",
    loginBtn: "登入",
    loginBusyMsg: "驗證中…",
    loginDefaultPwHint: "第一次使用？預設密碼為「Farspeed」，登入後可自行更改。",
    changePasswordLink: "更改密碼",
    changePasswordTitle: "更改密碼",
    logoutBtn: "登出",
    pwCurrentLabel: "目前密碼",
    pwNewLabel: "新密碼",
    pwConfirmLabel: "確認新密碼",
    pwTooShortMsg: "新密碼須至少4個字元。",
    pwMismatchMsg: "新密碼與確認密碼不相符。",
    pwCurrentWrongMsg: "目前密碼不正確。",
    pwChangedMsg: "密碼已更改，下次登入請使用新密碼。",
    pwSaveBtn: "儲存新密碼",
    signedInNone: "未登入",

    fRecordedBy: "記錄人",
    fRecordedByHint: "必填 — 由誰負責此次拆櫃／CFS／送貨",
    recordedByRequiredMsg: "儲存前請選擇記錄人。",
    saveBlockedMsg: (fields) => `儲存前請先填寫：${fields}。`,

    resetBtn: "重設所有送貨記錄（示範用）",
    resetConfirmMsg: "此操作將清除所有項目之送貨記錄，令所有貨物顯示為尚未送貨，但不會刪除任何倉存記錄。是否繼續？",
    resetDoneMsg: "所有送貨記錄已清除。",

    navJobLog: "單號記錄",
    navBilling: "帳單",
    navUpload: "上載",
    navIncoming: "待到倉",
    incomingTitle: "待到倉",
    incomingDesc: "已上載裝箱單但尚未辦理到倉手續的貨件。選擇哪些件號經拆櫃或CFS到倉及日期 \u2014 系統會建立或更新相應的存倉記錄。",
    incomingUploadToggle: "上載裝箱單",
    incomingUploadShow: "顯示",
    incomingUploadHide: "收起",
    incomingShowCompleted: "顯示已全部到倉的貨件",
    incomingNoneMsg: "目前沒有待到倉貨件 \u2014 上載裝箱單以新增。",
    incomingCaseCount: (n) => `裝箱單共 ${n} 件`,
    incomingLinkedTo: (id) => `已連結至 ${id}`,
    incomingFullyCheckedIn: "已全部到倉",
    incomingRemainingBadge: (remaining, total) => `尚有 ${remaining}／${total} 件未到倉`,
    incomingSelectCasesLabel: "選擇哪些件號到倉",
    incomingCheckInBtn: (n) => n > 0 ? `辦理到倉 (${n})` : "辦理到倉",
    incomingCheckedInNote: (incId) => `由待到倉記錄 ${incId} 辦理到倉`,
    packingListAddToIncomingBtn: (n) => `新增 ${n} 組至待到倉`,
    packingListShipmentCbmLabel: "整批CBM",
    packingListDistributeCbmBtn: "按重量分攤",
    packingListDistributeCbmHint: "將整份文件之總體積按各批次重量比例分攤。此為估算 \u2014 如有各批次實際數據請直接輸入。",
    dirInlineEditBtn: "編輯此地盤",
    dirInlineEditTitle: "目錄資料",
    dirInlineEditHint: "即時儲存至目錄，日後匯入此地盤將自動採用。",
    dirInlineSavedMsg: "已更新地盤資料。",
    packingListColProject: "項目／地盤",
    packingListAddSiteBtn: (code) => `將 ${code} 加入目錄`,
    packingListSitesLabel: (n) => `本次匯入涉及之地盤（${n}）`,
    packingListProjectFromCommon: "\u2014 使用上方欄位 \u2014",
    packingListProjectUnknown: (code) => `目錄中沒有 ${code} \u2014 請選擇地盤或先行新增。`,
    packingListMultiProjectHint: (n) => `此檔案涵蓋 ${n} 個項目，各批次會分別歸入其地盤。匯入前請檢查此欄；上方欄位只適用於未指定之批次。`,
    packingListCasesHint: "此批次之件號，以逗號分隔。更改名稱不影響各件重量及體積；增減件數則會改變批次內容，上方總數亦隨之更新。",
    packingListRemoveGroupBtn: "移除此組",
    incomingDeleteBtn: "刪除此待到倉貨件",
    incomingDeleteConfirm: "確定刪除此待到倉貨件？此操作只會移除待到倉記錄 \u2014 已辦理到倉的存倉記錄不受影響。",
    packingListIncomingHint: "此步驟只會將件號加入待到倉名單 \u2014 倉庫、拆櫃/CFS、工單號及日期會在稍後於「待到倉」分頁辦理到倉時才決定。",
    billingTitle: "存倉收費",
    billingDesc: "可按客戶、項目、工單號或件號搜尋存倉收費。收費以到倉批次計算（不按地盤分開），按客戶的每CBM月費計：先扣免費日數，該月餘下日數按日比例計算，其後每個月則全額收取——即使貨物於月中送出亦然。",
    billingSearchPlaceholder: "客戶、項目、工單號或件號",
    billingColClient: "客戶",
    billingColProject: "項目／地盤",
    billingColJobNo: "工單號",
    billingColBatchDate: "到倉日期",
    billingColCbm: "CBM",
    billingColRate: "收費",
    billingColStatus: "狀態",
    billingColTotal: "總額",
    billingNoneMsg: "此搜尋未有可收費的存倉記錄。",
    billingPerCbmMonth: "CBM/月",
    billingOngoing: "計算中",
    billingClosed: "已送出",
    billingShowBtn: "顯示明細",
    billingHideBtn: "收起",
    billingCasesLabel: "件號",
    billingFreeDaysNote: (d) => `已扣除${d}日免費存倉。`,
    billingEstimatedNote: "＊按件數比例估算已送/餘下CBM分配（此記錄沒有逐件CBM資料）。",
    billingGrandTotal: "總計",
    billingFootnote: "計算中的項目按至今日計算，直至貨物標記為已送出才會停止累加。收費可於 目錄 → CBM 收費 設定。",
    billingSelectAllAria: "選取畫面所有記錄",
    billingSelectedCount: (entries, rows, total) => `已選 ${entries} 項記錄 \u2014 ${rows} 行收費，共 ${total}`,
    billingDeleteSelectedBtn: (n) => `刪除 ${n} 項記錄`,
    billingDeleteSelectedHint: "整批只需輸入一次密碼。刪除記錄會一併移除其所有收費行、到倉及送貨記錄。",
    billingDeleteItemBtn: "刪除此記錄（需要管理密碼）",
    adminConfirmTitle: "確認管理員操作",
    adminConfirmDesc: (name) => `請重新輸入 ${name || "您"} 的密碼以繼續。此操作會永久刪除相關的存倉記錄，不只是這一行帳單。`,
    adminConfirmBtn: "永久刪除",
    billingModeSearch: "搜尋",
    billingModeMonthly: "每月總覽",
    billingModeHandling: "拆櫃/送貨收費",
    billingHandlingDesc: "按客戶收費表計算的拆櫃/送貨服務費及貨櫃拖運費 \u2014 與上方的倉租計算分開。只顯示已有收費表的客戶（其士、迅達）。",
    billingHandlingNeedsQuote: (n) => `${n} 項工作需要人手報價 \u2014 收費表未有提供該地區／貨物類型組合的自動收費。`,
    billingHandlingNoneMsg: "暫無拆櫃/送貨收費 \u2014 當已有收費表的客戶之貨品記錄到倉或送貨後，將會顯示於此。",
    billingHandlingColType: "工作類型",
    billingHandlingColBasis: "計算基礎",
    billingHandlingColRate: "費率",
    billingHandlingTypeDevan: "拆櫃",
    billingHandlingTypeDelivery: "送貨",
    billingHandlingTypeHaulage: "貨櫃拖運",
    billingHandlingOversizeTag: (mult) => `超大件 \u00d7${mult}`,
    billingHandlingHaulageBasis: (c20, c40) => [c20 ? `${c20} \u00d7 20呎` : "", c40 ? `${c40} \u00d7 40呎` : ""].filter(Boolean).join("，"),
    billingHandlingQuoteBadge: "另行報價",
    billingHandlingFootnote: "R/Ton（運費噸）= max(重量噸數, CBM體積)。費率來自2018年收費表；地區及貨物類型於每項記錄設定（於手動輸入編輯）。超大件倍數與工作紙上顯示的相同級別一致。",
    billingModeInvoices: "發票核對",
    invoicesTitle: "客戶發票 \u2014 與系統記錄核對",
    invoicesDesc: "記錄發給客戶之每張發票，系統會按同一客戶、地盤、類別及月份計算應收金額並作核對，如有差異會即時顯示。",
    invoicesNoneMsg: "此月份未有發票記錄。",
    invoicesFootnote: "涵蓋存倉、CFS、拆櫃及送貨。吊運及搬移暫未包括。點擊該行可查看系統金額之明細。",
    invoiceAddBtn: "新增發票",
    invoiceColNo: "發票編號",
    invoiceColDate: "日期",
    invoiceColSite: "地盤",
    invoiceColCategory: "類別",
    invoiceColInvoiced: "發票金額",
    invoiceColExpected: "系統金額",
    invoiceColDifference: "差異",
    invoiceStatusMatch: "相符",
    invoiceStatusOver: "多收",
    invoiceStatusUnder: "少收",
    invoiceStatusNothing: "無相關收費",
    invoiceNothingHint: "系統於此客戶、地盤、類別及月份沒有任何應收項目。請檢查地盤名稱是否相符，以及到倉／送貨是否已記錄。",
    invoiceLinesLabel: (n) => `系統金額由 ${n} 項收費組成`,
    invoiceEstimatedTag: "估算",
    invoiceTotalInvoiced: "發票總額",
    invoiceTotalExpected: "系統總額",
    invoiceTotalDifference: "差異",
    invoiceProblemCount: (n) => `${n} 張發票需核對`,
    invoiceUninvoicedLabel: "本月有收費但未有發票記錄",
    invoiceViewScanBtn: "檢視掃描檔",
    invoiceDueDateLabel: "付款到期日",
    invoiceAmountLabel: "金額 (HK$)",
    invoiceOrderedByLabel: "落單人",
    invoiceOrderRefLabel: "訂單參考",
    invoiceChargeLineLabel: "收費說明",
    invoiceChargeLineHint: "顯示於發票之字句，例如 CHARGES FOR JANUARY 2026 AT LUMPSUM",
    invoiceNarrativeLabel: "附加說明（可選）",
    invoiceBillToLabel: "發票地址",
    invoiceRevisedDateLabel: "修訂日期",
    invoiceRevisedByLabel: "修訂人",
    invoiceScanLabel: "上載掃描檔",
    invoiceAttachingMsg: "正在上載掃描檔⋯",
    invoicePreviewMsg: (expected, diff) => `系統於此客戶、地盤、類別及月份之金額為 ${expected}，差異 ${diff}。`,
    debitNotePrintLabel: (no) => `發票 ${no}`,
    billingMonthLabel: "月份",
    billingYearLabel: "年份",
    billingMonthNoneMsg: "此月份沒有存倉收費。",
    billingMonthFootnote: "各客戶總額應與其MYOB該月發票金額相符，可用作出單前核對。點擊客戶可查看組成總額的每一項明細。",
    jobLogSearchPlaceholder: "工單號、地盤、地盤代號、梯號、SHK⋯",
    jobLogFromLabel: "由",
    jobLogToLabel: "至",
    jobLogNoMatchMsg: "沒有符合篩選條件的工單。",
    jobLogCount: (shown, total, jobs) => shown === total
      ? `共 ${total} 張工單 \u00b7 ${jobs} 個單號`
      : `${total} 張工單中的 ${shown} 張 \u00b7 ${jobs} 個單號`,
    jobLogTitle: "所有已使用的工作單號",
    jobLogDesc: "所有曾建立的Devan、CFS及送貨單號，最新在前。點擊任何一行可查看及重印該工單。",
    jobLogColJobNo: "單號",
    jobLogColType: "類型",
    jobLogColDate: "日期",
    jobLogColClient: "客戶",
    jobLogColSite: "項目/地盤",
    jobLogColRecordedBy: "記錄人",
    jobLogNoneMsg: "尚未建立任何單號。",
    viewReprintBtn: "查看／重印",

    navCancelledJobs: "已取消工作",
    cancelJobBtn: "取消",
    cancelledJobsTitle: "已取消工作",
    cancelledJobsDesc: "已作廢之Devan、CFS及送貨工單。此類記錄不會出現在主要之「單號記錄」，並保留其單號以免被重用。由於此為示範用途，可於此處永久刪除以清理測試資料。",
    cancelledJobsNoneMsg: "沒有已取消之記錄。",
    restoreBtn: "還原",
    purgeBtn: "永久刪除",
    permanentDeleteConfirmMsg: "此操作將永久刪除該記錄並釋放其單號，無法還原。是否繼續？",

    inventoryRemainingLabel: "存倉中",
    inventoryNoRemainingPkgsMsg: "所有件號均已送出，倉內已無剩餘。",

    settingsLabel: "設定",
    navDuplicatesShort: "重複記錄",
    navCheckIns: "到倉核對",
    navCheckInsCount: (n) => `到倉核對（多 ${n}）`,
    fAwaitingCollection: "尚無到倉日期 \u2014 貨物仍待提貨",
    legacyFieldNames: { client: "客戶", site: "地盤", date: "日期" },
    legacyRowMissingHint: (list) => `此檔案需填寫${list}方可處理。日期決定其於存倉紀錄中之排序及起算收費時間，不可留空。`,
    navPlReader: "裝箱單讀取",
    plrTitle: "裝箱單讀取",
    plrDesc: "一次讀取多份裝箱單 \u2014 Excel 直接讀取，PDF 則交由與單檔畫面相同之掃描器處理。結果可匯出為一份表格，再由「裝箱單匯入」一次載入。使用前請先核對最後一欄：凡件數與件號不符者均會標示。",
    plrDupCase: (code, files) => `件號 ${code} 在同一參考編號下重複出現（${files}）`,
    plrRefClientClash: (list) => `此參考編號對應兩個不同客戶：${list}`,
    plrRefSiteClash: (list) => `此參考編號對應兩個不同地盤：${list}`,
    plrChooseBtn: "選擇裝箱單…",
    plrReading: (name) => `正在讀取 ${name}…`,
    plrSaveBtn: "儲存表格",
    plrSavedNote: (n) => `已儲存 ${n} 行`,
    plrSaveFailed: (err) => `無法儲存：${err}`,
    plrRestored: (n, at) => `已從 ${at || "上次"} 回復 ${n} 行`,
    plrClearConfirm: (n) => `確認清除全部 ${n} 行？\n\n尚未匯出之修改將會遺失。`,
    plrExportBtn: "匯出表格",
    plrClearBtn: "重新開始",
    plrTooBig: (kb) => `檔案過大（${kb} KB），無法掃描 \u2014 請分割或改用 Excel 版本`,
    plrCount: (rows, files, off) => `${files} 份檔案共 ${rows} 個批次${off ? ` \u00b7 其中 ${off} 個件數與件號不符` : ""}`,
    navLedger: "存倉紀錄",
    ledgerTitle: "存倉紀錄",
    ledgerDesc: "依地盤列出所有到倉及送貨記錄，按日期排列並附累計餘數 \u2014 格式與客戶存倉清單一致，便於對照。此處不列載費用及免費期，該等項目依據收費規則而非出入記錄。",
    legacyReturnedNote: (job) => `貨箱於工單 ${job} 退回`,
    ledgerNoDate: "未填日期",
    ledgerAddDateHint: "輸入此筆出入倉之實際日期 \u2014 完成後即會計入餘數",
    ledgerNotCounted: "不計入",
    ledgerUndatedTotal: (n) => `共 ${n} 筆未填日期 \u2014 已於下方地盤以紅色標示，在補上日期前不計入任何餘數。`,
    ledgerAsOfLabel: "截至日期",
    ledgerAsOfToday: "今日",
    ledgerTotalOnHand: (d) => `件存倉（${d ? `截至 ${d}` : ""}），全部地盤合計`,
    ledgerSetAside: (later, undated) => [
      later ? `${later} 筆日期晚於此日` : "",
      undated ? `${undated} 筆未填日期` : "",
    ].filter(Boolean).join(" \u00b7 ") + "，不計入餘數。",
    ledgerExportBtn: "匯出紀錄",
    ledgerSiteLine: (i, o, b) => `入 ${i} \u00b7 出 ${o} \u00b7 現存 ${b}`,
    ledgerColDate: "日期",
    ledgerColJob: "FC 工單號",
    ledgerColDir: "出 / 入",
    ledgerColType: "類型",
    ledgerColIn: "入倉件數",
    ledgerColOut: "出倉件數",
    ledgerColBalance: "累計餘數",
    storeListTitle: "與客戶存倉紀錄核對",
    storeListDesc: "載入客戶之存倉清單（每個地盤一頁，每筆出入倉一行）。系統會以 FC 工單號逐筆核對，僅列出不符之項目。若紀錄未包含本月資料，本月工單將顯示為差異，屬正常情況。",
    storeListChooseBtn: "選擇存倉清單…",
    storeListLoaded: (name, n, sites) => `${name} \u2014 ${n} 筆出入記錄，共 ${sites} 個地盤`,
    storeListUnreadable: (err) => `無法讀取此檔案：${err}`,
    storeListSiteLine: (i, o, l, ai, ao, al) => `紀錄：入 ${i}、出 ${o}、餘 ${l} \u00b7 倉庫：入 ${ai}、出 ${ao}、餘 ${al}`,
    storeListSiteAgrees: (n) => `全部 ${n} 筆記錄均與倉庫相符。`,
    storeListColJob: "FC 工單號",
    storeListColDir: "出 / 入",
    storeListColDm: "紀錄之 DM",
    storeListColSheet: "紀錄",
    storeListColApp: "倉庫",
    storeListColDiff: "差異",
    checkInsTitle: "到倉記錄與單據核對",
    checkInsDesc: "列出所有到倉記錄，以其參考編號分組。參考編號本身已註明件數，若其下各筆合計超出該數，即同一批貨被重複到倉 \u2014 通常是整批一次、再按電梯分批一次。",
    checkInsOverBanner: (n) => `共有 ${n} 件被重複計算。還原到倉記錄會將貨箱退回「待到倉」，可重新妥善辦理。`,
    checkInsHeldVsStated: (held, stated, left) => `已到倉 ${held} 件，單據註明 ${stated} 件 \u00b7 現存 ${left} 件`,
    checkInsHeldOnly: (held, left) => `已到倉 ${held} 件 \u00b7 現存 ${left} 件 \u00b7 參考編號未註明件數，無法核對`,
    checkInsOverBy: (n) => `多出 ${n} 件`,
    checkInsRowCount: (n) => `${n} 筆到倉`,
    checkInsSearchLabel: "搜尋到倉記錄",
    checkInsSearchPlaceholder: "記錄、貨件、電梯或參考編號 \u2014 例如 INC-0559",
    checkInsColEntry: "記錄",
    checkInsColCode: "電梯 / 訂單",
    checkInsColArrived: "到倉日期",
    checkInsColUnits: "件數",
    checkInsColLeft: "現存",
    checkInsColSource: "來源",
    checkInsReverseBtn: "還原此到倉記錄",
    checkInsMergeLabel: "合併至…",
    checkInsMergedNote: (from) => `由 ${from} 轉入`,
    checkInsMergeConfirm: (from, keep, n, dels) => `確認將 ${from} 合併至 ${keep}？\n\n${dels ? `其 ${dels} 筆送貨記錄會先轉至 ${keep}，不會遺失。` : ""}隨後還原 ${from} 之到倉記錄，其 ${n} 件貨箱退回「待到倉」。\n\n此操作無法從此頁還原。`,
    checkInsReverseHasDeliveries: (id, n) => `請注意 \u2014 ${id} 已有 ${n} 筆送貨記錄。\n\n還原其唯一到倉記錄後，該筆送貨將無對應存貨。如此筆為重複記錄，請先記下其送貨詳情，以便轉記於保留之記錄。\n\n是否繼續？`,
    checkInsReverseConfirm: (id, n, gone) => `確認還原此到倉記錄？\n\n${n} 件貨箱將退回「待到倉」。${gone ? `\n\n${id} 已無其他記錄，會一併刪除。` : ""}\n\n此操作無法從此頁還原。`,

    depotOverviewTitle: "各倉存倉概覽",
    depotOverviewItemsLabel: "項",

    newEntryManual: "手動輸入",
    newEntryImport: "匯入",

    moveCasesLabel: "將貨箱移至其他記錄",
    moveCasesHint: "適用於誤配梯號之貨箱。選取貨箱及目標記錄，貨箱連同到倉批次一併轉移。即時生效，兩邊記錄同時更新。已送出之貨箱不會顯示，因其已屬送貨記錄。",
    moveCasesDestPlaceholder: "選擇目標記錄⋯",
    moveCasesBtn: "移動貨箱",
    moveCasesDoneMsg: (codes, dest) => `已將第 ${codes} 件移至 ${dest}。`,
    splitArrivalLabel: "分批到倉（貨箱不同日子到貨）",
    splitArrivalHint: "如同一張裝箱單不是同一日全部拆櫃／送到倉，可在此分批記錄，每批有自己的日期及類型。抵倉日期會自動跟隨最早一批。",
    splitArrivalCasesCol: "貨箱",
    splitArrivalAddBtn: "新增到倉批次",
    splitArrivalSelectHint: "點選此批次到達的貨箱，然後新增。",
    splitArrivalAllAssignedMsg: "所有貨箱已分配到倉批次。",
    badgePartialArrival: (a, b) => `已到倉 ${a}/${b} · 其餘未到`,
    notYetArrivedTag: "未到倉",
    notYetArrivedHint: "此貨箱尚未到倉，未能安排送出。",
    pendingArrivalNotice: (n) => `${n} 箱尚未到倉，未能選擇送貨。貨到後請在編輯頁記錄其到倉批次。`,

    tabFreeStorage: "免費存倉",
    freeStorageTitle: "免費存倉日數",
    freeStorageDesc: "標準免費存倉為每批拆櫃／CFS到倉日起計14日（Schindler為21日）。如個別項目另有安排（例如Otis按項目而定），可在下方新增規則：當項目或地盤名稱包含所輸入的文字時，即採用指定日數，並優先於客戶預設。",
    fFreeProject: "項目名稱包含",
    fFreeProjectHint: "例如：Kwu Tung North Area 19",
    fFreeDays: "免費日數",
    freeStorageAddBtn: "新增規則",
    freeColProject: "項目條件",
    freeColDays: "免費日數",
    freeStorageNoneMsg: "未有項目規則 — 採用標準14日（Schindler 21日）。",
    tabPricing: "CBM 收費",
    pricingTitle: "各客戶 CBM 收費",
    pricingDesc: "貨倉及露天場地存倉均以CBM計算收費，此處設定每CBM的收費；KG只用於計算送貨及CFS費用。留空即採用預設收費（顯示於欄位提示中）。",
    pricingColClient: "客戶",
    pricingColRate: "收費",
    pricingPerCbm: "／ CBM",
    pricingResetBtn: "回復預設",
    tabLegacy: "舊資料上載",
    legacyUploadTitle: "舊資料上載",
    legacyUploadDesc: "批量上載舊有的拆櫃、CFS、送貨、調位、吊運等工單檔案。建議分兩輪上載：先上載所有拆櫃/CFS檔案（建立到倉記錄），再上載送貨檔案 — 系統會自動配對並完結對應的到倉記錄。不論類型，所有檔案均會存檔並可於下方搜尋。",
    legacyChooseFilesBtn: "選擇檔案…",
    legacyDocType: "工單類型",
    legacyProjectSite: "項目／地盤",
    legacyUnitCode: "單位／升降機編號",
    legacyPkgs: "件數",
    legacyWeightKg: "重量 (kg)",
    legacyCbm: "CBM",
    legacyAlreadyDelivered: "已送出 — 直接完結此工作，不留在倉內",
    legacyProcessBtn: (n) => `處理 ${n} 個檔案`,
    legacyProcessingMsg: "處理中…",
    legacyResultsMsg: (archived, created, delivered, enriched, checkedIn) => `已存檔 ${archived} 個檔案${created > 0 ? `，並建立 ${created} 項存倉記錄` : ""}${checkedIn > 0 ? `，並為 ${checkedIn} 組件號辦理待到倉核對` : ""}${enriched > 0 ? `，並配對 ${enriched} 項現有記錄並補充其資料` : ""}${delivered > 0 ? `，並為 ${delivered} 項已到倉記錄登記送貨` : ""}。`,
    legacyImportedNote: (name) => `由舊資料檔案匯入：${name}`,
    legacyAutoClosedNote: "舊資料匯入時自動完結（標記為已送出）。",
    legacyBacklogTitle: "待處理記錄",
    legacyBacklogDesc: "所有透過舊資料上載的檔案，不論是否建立了存倉記錄。",
    legacyBacklogNoneMsg: "尚未上載任何舊資料檔案。",
    legacyBacklogNoMatchMsg: "沒有符合篩選條件的檔案。",
    legacyBacklogSortLabel: "排序",
    legacyBacklogSortRecent: "最新在前",
    legacyBacklogSortJobNo: "按工單號",
    legacyBacklogCount: (shown, total) => shown === total ? `共 ${total} 個檔案` : `${total} 個檔案中的 ${shown} 個`,
    legacyColFile: "檔案",
    legacyColLinked: "連結記錄",
    legacyArchivedOnly: "僅存檔",
    legacyEditLinkedHint: "上方欄位修正此存檔記錄。下方為此檔案所建立之倉存記錄 \u2014 於此修改會直接更新相應之 FS-#### 記錄。",
    legacyLinkedRecordsLabel: "此檔案建立之記錄",
    legacyLinkedRecordsHint: "於此輸入之重量或體積會視為該批次之實際數據，並取代由單據讀取之數值。兩者留空則回復使用裝箱單數據。更改「記錄批次」會將整筆記錄轉至該批次，適用於工單原先入錯電梯之情況。如需更改所涉貨箱，請於存倉列表開啟該記錄。",
    legacyRecordEntryLabel: "記錄批次",
    legacyDeleteTitle: "由存檔中刪除此檔案。如選擇還原，亦會撤銷此檔案對倉存所作之更改：",
    legacyDeleteArrivalLine: (label, n, cases) => `${label}：撤銷 ${n} 筆到倉記錄，涉及 ${cases} 個貨箱。`,
    legacyDeleteDeliveryLine: (label, n, cases) => `${label}：撤銷 ${n} 筆送貨記錄，涉及 ${cases} 個貨箱 \u2014 該等貨箱回復存倉。`,
    legacyDeleteEntryRemoved: (label) => `${label} 由此檔案建立，且再無其他記錄，將一併刪除。`,
    legacyDeleteEntryKept: (arrivals, deliveries) => `該批次保留 \u2014 尚有 ${arrivals} 筆到倉及 ${deliveries} 筆送貨記錄。`,
    legacyDeleteIncomingLine: (incId, n) => `待到倉記錄 ${incId}：${n} 個貨箱回復待辦到倉。`,
    legacyDeleteBlockedMsg: (label, n) => `${label} 由此檔案建立，但其後已有 ${n} 筆送貨記錄，故予以保留。如需刪除，請先還原該等送貨記錄。`,
    legacyDeleteStrandedMsg: (labels) => `${labels} 已無此檔案所建立之記錄，故不作處理。`,
    legacyDeleteNothingMsg: "此檔案並未建立任何倉存記錄，無需還原 \u2014 只會刪除存檔記錄。",
    legacyDeleteHint: "此檔案為現有批次填入之地盤名稱、提單資料等不會撤銷。回復待辦之貨箱可於「待到倉」重新辦理到倉。",
    invSelectedCount: (n, pkgs) => `已選 ${n} 筆 \u00b7 ${pkgs} 件`,
    invBulkDeleteBtn: "刪除已選",
    invBulkDeleteConfirm: (n, pkgs, delivered) => `確認永久刪除 ${n} 筆倉庫記錄及其 ${pkgs} 件貨物？\n\n${delivered ? `其中 ${delivered} 筆已有送貨記錄，一併刪除。\n\n` : ""}如尚未匯出資料，請先匯出。此操作無法還原。`,
    incomingSelectAll: (n) => `全選目前 ${n} 筆`,
    incomingSelectedCount: (n, cases) => `已選 ${n} 筆 \u00b7 ${cases} 個貨箱`,
    incomingBulkDeleteBtn: "刪除已選",
    incomingBulkDeleteConfirm: (n, cases, linked) => `確認刪除 ${n} 筆待到倉記錄及其 ${cases} 個貨箱？\n\n${linked ? `其中 ${linked} 筆已到倉至倉庫記錄 \u2014 該等記錄會保留，但將無裝箱單依據。\n\n` : ""}此操作無法還原。`,
    legacySelectAllHint: "選取目前篩選結果中所有檔案",
    legacySelectedCount: (n) => `已選 ${n} 份檔案`,
    legacyBulkReverseBtn: "還原記錄並刪除",
    legacyBulkDeleteBtn: "僅刪除存檔記錄",
    legacyClearSelection: "取消選取",
    legacyBulkDeleteConfirm: (n) => `確認刪除 ${n} 份存檔記錄？\n\n其所建立之倉庫記錄將完全保留，僅刪除存檔。\n\n此操作無法還原。`,
    legacyBulkReverseConfirm: (n, entries, removed) => `確認還原並刪除 ${n} 份檔案？\n\n將影響 ${entries} 筆倉庫記錄，其中 ${removed} 筆會被刪除。貨箱退回「待到倉」。\n\n執行前請先匯出倉庫資料。此操作無法還原。`,
    legacyDeleteReverseBtn: "還原記錄並刪除檔案",
    legacyDeleteKeepRecordsBtn: "只刪除檔案，保留記錄",
    legacyDeleteListingOnlyBtn: "刪除檔案",
    legacyRecordMoveNote: (from, to, cases, kept) => `儲存後由 ${from} 轉至 ${to}。${cases ? (kept === cases ? `全部 ${cases} 個貨箱一併轉移。` : `${to} 只有 ${cases} 個箱號中之 ${kept} 個，其餘 ${cases - kept} 個將被移除，請核對上方重量及體積。`) : ""}`,
    legacyClientUnresolved: "— 請選擇客戶 —",
    legacyClientRequiredSummaryMsg: "部分檔案未能識別客戶 — 處理前請為每個檔案選擇正確客戶。",
    legacyDeliveredFrom: (id) => `送出自 ${id}`,
    legacyUnmatchedReferral: (job) => `找不到工單號 ${job} 的到倉記錄`,
    legacyUnmatchedHint: "此送貨記錄指向的工單號尚未存在系統內 — 請先上載其拆櫃/CFS檔案，或稍後手動連結。",
    legacyProjectSiteEn: "項目／地盤（英文）",
    legacyProjectSiteZh: "項目／地盤（中文）",
    legacySiteRequiredMsg: "請至少填寫英文或中文地盤名稱其中一項。",
    legacySiteRequiredSummaryMsg: "部分檔案的地盤名稱（中英文）均未填寫 — 處理前請至少填寫其中一種語言。",
    legacyScanningMsg: "讀取檔案中…",
    legacyAutoDetectHint: "Excel檔案 (.xlsx/.xls/.csv) 直接讀取，PDF則以掃描方式讀取 — 客戶、地盤、工單號、日期、提單資料、參照工單、件號及總數均會自動填入。請於處理前檢查下方欄位；掃描結果不及試算表準確。圖片檔仍需手動輸入。",
    legacyAutoDetectedTag: "已從檔案自動偵測 — 請核對",
    legacyReferLine: (job, date) => `指向工單號 ${job}，日期 ${date}`,
    legacyReferJobNoLabel: "指向到倉工單號",
    legacyEnrichedNote: (name) => `由舊資料檔案補充資料：${name}`,
    legacyMatchedIncoming: (id) => `已配對至待到倉 ${id} \u2014 請選擇此檔案中已到達的件號`,
    legacyNoPackingListTitle: "本地盤未有裝箱單記錄",
    legacyPackingListLookedFor: (client, site) => `已按客戶「${client || "（未填）"}」及地盤「${site || "（未填）"}」搜尋待到倉貨件，未有結果。`,
    legacyPackingListNearMiss: (n, list) => `另有 ${n} 筆待到倉貨件登記於其他名稱之下 \u2014 ${list}。如其中一筆即為本批貨，請先修正上方之客戶或地盤名稱，切勿在此另建裝箱單。`,
    legacyNoPackingListDesc: "系統沒有可供入倉的貨件。此工單本身已列明批次及件號，可直接據此建立裝箱單 \u2014 該批次之重量及體積會平均分配至各件（工單只提供批次總數）。建立後可於待到倉逐件修改。",
    legacyCreatePackingListBtn: (n) => `建立裝箱單 \u2014 ${n} 個批次`,
    legacyPackingListFromSheetNote: (src) => `由工單 ${src} 手動建立之裝箱單`,
    legacyMatchedIncomingCount: (n) => `此地盤配對到 ${n} 項待到倉貨件 \u2014 請分別選擇此檔案中已到達的件號`,
    legacyReferJobNoHint: "可選填 \u2014 填寫後只會配對該工單號的到倉記錄；留空則按客戶＋地盤配對（可能配對多項）。",
    legacyOversizeLabel: "超大件",
    legacyOversizeCbmPlaceholder: "CBM",
    legacyOversizeHint: "將按此客戶的超大件收費標準計算，以此CBM為基礎。",
    legacyMatchedItemsCount: (n) => `配對到 ${n} 項倉內仍有貨件的存倉記錄 \u2014 請選擇此次送貨送出的件號`,
    legacyMatchedItemsReturn: (n) => `配對到 ${n} 項已出貨之存倉記錄 \u2014 請選擇本次退回之件號`,
    legacyMatchedItemReturn: (id) => `退回至 ${id}`,
    legacyTypeSelectPlaceholder: "例如 1,3-5,7",
    legacyTypeSelectBtn: "加入",
    legacySelectedTotals: (count, kg, cbm) => `已選：${count} 件 \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacySelectedTotalsGrand: (count, kg, cbm) => `全部已選（合計）：${count} 件 \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacyDeliveryDeclaredLabel: "此送貨單據數據：",
    legacyOversizeCasesPh: "例如 14/21",
    legacyOversizeCaseCol: "件號",
    legacyOversizeAdd: "+ 新增超大件",
    legacyOversizeRemove: "移除",
    legacyOversizeTotal: (n, cbm) => `${n} 件超大件 \u00b7 合計 ${cbm} cbm \u00b7 各按其級距計費`,
    legacyOversizeDeliveryHint: "會列印於此送貨單的「超大件」欄",
    legacyDeliveryDeclaredGap: (declaredKg, listedKg, pct, heavier) =>
      `單據列明 ${declaredKg} kg；所選件號按裝箱單合計 ${listedKg} kg（${heavier ? "+" : "\u2212"}${pct}%）。記錄以單據為準。`,
    legacySheetTotalLabel: (lots, ref) => `單據總數${ref ? ` \u00b7 ${ref}` : ""} \u2014 涵蓋 ${lots}`,
    legacySheetTotalJob: (job) => `工單 ${job}`,
    legacySheetTotalHint: "單據就這幾個梯號只列一個總數。請在此修改，下方各梯號的分攤額會按裝箱單比例自動計算。",
    legacyDeclaredShareNote: (kg, cbm) => `分攤自單據總數：${kg} kg \u00b7 ${cbm} cbm`,
    incomingDeclaredLabel: "拆櫃／CFS 單據總數（可留空）",
    legacyDeclaredLabel: "此單據數據：",
    legacyDeclaredHint: "倉租、裝卸及收費以單據數據為準；裝箱單重量仍按件保留記錄。",
    legacyDeclaredKgGap: (kg, heavier, pct) => heavier
      ? `單據比裝箱單重 ${kg} kg（+${pct}%）\u2014 屬正常，單據已包含桁架段的木箱及卡板。`
      : `單據比裝箱單輕 ${kg} kg（\u2212${pct}%）\u2014 處理前請先核對。`,
    legacyDeclaredPkgsGap: (declared, selected) => `單據件數為 ${declared} 件，現選 ${selected} 件。`,
    fWeightFromSheet: "取自拆櫃／CFS 單據",
    fWeightSheetShare: "估算 \u2014 分攤自涵蓋多個梯號的單據總數",
    fWeightPackingList: (kg) => `裝箱單：${kg} kg`,
    fVolumePackingList: (cbm) => `裝箱單：${cbm} cbm`,
    fVolumeFromOversize: "超大件 CBM \u2014 收費級距按此計算，優先於單據總數",
    legacyMatchedItem: (id) => `送出自 ${id}`,
    legacyArrivalStaysOpenHint: "此記錄會保持在倉狀態，直至上載對應的送貨檔案（或日後手動記錄送貨）為止。",
    legacyNoReferralHint: "未有偵測到「Ref Job no.」字句 — 請手動輸入到倉工單號，否則此檔案只會被存檔。",
    legacySheetCasesNote: (mark, n) => `工單註明 ${mark} \u2014 已預先選取 ${n} 件`,
    legacyCasesFormatClash: (sheet, list) => `全部不符，故未能預先選取。工單要求 ${sheet}…而裝箱單記載為 ${list}… \u2014 同一批貨箱但寫法不同。請重新匯入裝箱單以重建件號，或今次先手動選取。`,
    legacySheetCasesMissing: (list) => `第 ${list} 件不在倉內`,
    legacySheetCasesElsewhere: (list, more) => `但已入於其他批次：${list}${more ? `，另有 ${more} 件` : ""} \u2014 請改由該批次出貨，或於存倉列表調整貨箱歸屬`,
    legacyReplaceCasesHint: (n) => `件數相同，但其中 ${n} 個件號與本工單不符。兩者皆可能有誤 \u2014 掃描可能漏字，工單亦可能手誤 \u2014 請先核對紙本再作更改：`,
    legacyReplaceCasesBtn: "將此批到貨件號改為工單所示",
    legacyScannedFromPdf: "此為掃描讀取結果 \u2014 處理前請與紙本核對各欄位及件號。",
    legacyScanFailed: (why) => `無法讀取此PDF（${why}）。請手動輸入，或改為上載Excel原檔。`,
    legacyCasesArePicture: "此單之件號為貼上之圖片而非文字，無法讀取。請於下方件號欄輸入，或填入匯入表之 Case Numbers 欄。",
    legacyCaseCountMismatch: (lot, stated, listed) => `${lot}：工單註明 ${stated} 件，但只列出 ${listed} 個件號 \u2014 處理前請核對 C/S No. 清單。`,
    legacyIncomingCasesMissing: (list) => `第 ${list} 件不在此批到貨內`,
    legacyCasesFoundInMore: (n) => `\u2026另有 ${n} 批到貨屬同一批次。`,
    legacySheetCasesAmbiguous: (mark, ids) => `工單註明 ${mark}，但 ${ids} 均符合此批次 \u2014 未有預先選取，請自行選擇貨箱。`,
    legacyCasesFoundIn: (codes, id, unit) => `\u2192 ${id} \u00b7 ${unit} 以相同訂單號載有第 ${codes} 件。僅憑件號不足以確定為同一箱 \u2014 請先核對文件再於該項辦理到倉。`,
    legacyMisfiledCases: (codes, from) => `第 ${codes} 件現存於 ${from}，但件號所示之總件數顯示應屬此項 \u2014 裝箱單之梯號欄有誤。`,
    legacyMisfiledFixBtn: "移至此項",
    legacySomeArrivalsUnmatched: (jobs) => `找不到到倉工單號 ${jobs} 的存倉記錄 \u2014 此工單該部分不會記錄送貨。請先上載該拆櫃/CFS檔案，或修正上方號碼。`,
    legacyArrivalLaterInBatch: (job, file) => `工單號 ${job} 之到倉檔案已在本批次內（${file}），尚未處理。按「處理」時會先建立到倉記錄，故此送貨單屆時可自行配對，無需另行處理。`,
    legacyNoArrivalFoundHint: (job) => `找不到工單號 ${job} 的存倉記錄 \u2014 請檢查號碼是否正確，或先上載該拆櫃/CFS檔案。此檔案仍會被存檔，但不會記錄任何送貨。`,
  },
};

function inputStyleFor(colors) {
  return {
    fontFamily: FONT_BODY,
    border: `1px solid ${colors.line}`,
    background: colors.surface,
    color: colors.ink,
  };
}
const inputClass = "px-2.5 py-1.5 rounded text-sm outline-none focus:ring-2";

function Badge({ children, tone = "grey", colors }) {
  const map = {
    grey: { bg: colors.surfaceDim, fg: colors.inkFaint },
    amber: { bg: colors.amberSoft, fg: colors.amberText },
    green: { bg: colors.greenSoft, fg: colors.green },
    red: { bg: colors.redSoft, fg: colors.red },
    navy: { bg: colors.navySoft, fg: colors.onDark },
  };
  const c = map[tone];
  return (
    <span
      className="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold tracking-wide"
      style={{ background: c.bg, color: c.fg, fontFamily: FONT_DISPLAY }}
    >
      {children}
    </span>
  );
}

function StatusBadge({ item, colors, t }) {
  const status = deriveStatus(item);
  if (status === "pending_collection") {
    const alert = lfdAlert(item);
    if (alert && alert.level === "overdue") return <Badge tone="red" colors={colors}>{t.badgeLfdOverdue(Math.abs(alert.days))}</Badge>;
    if (alert && alert.level === "soon") return <Badge tone="amber" colors={colors}>{t.badgeLfdSoon(alert.days)}</Badge>;
    return <Badge tone="grey" colors={colors}>{t.badgePendingCollection}</Badge>;
  }
  if (status === "at_depot") {
    const info = storageInfo(item);
    if (info.billable) return <Badge tone="red" colors={colors}>{t.badgeBillable(info.billableDays)}</Badge>;
    const pendingArr = notYetArrivedPackages(item).length;
    if (pendingArr > 0) {
      const arrived = (item.packages || []).length - pendingArr;
      return <Badge tone="amber" colors={colors}>{t.badgePartialArrival(arrived, (item.packages || []).length)}</Badge>;
    }
    const left = info.daysLeft != null ? info.daysLeft : info.freeDays - info.daysHeld;
    return <Badge tone="green" colors={colors}>{t.badgeFree(left)}</Badge>;
  }
  if (status === "partial") {
    const info = storageInfo(item);
    const del = deliveredUnits(item);
    const tot = totalUnits(item);
    return (
      <Badge tone={info.billable ? "red" : "amber"} colors={colors}>
        {t.badgePartial(del, tot, info.billable ? t.badgeBillableSuffix(info.billableDays) : "")}
      </Badge>
    );
  }
  const info = storageInfo(item);
  return <Badge tone="navy" colors={colors}>{t.badgeDelivered(info.billable ? t.badgeBilledSuffix(info.billableDays) : "")}</Badge>;
}

function Field({ label, children, hint, colors }) {
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
        {label}
      </span>
      {children}
      {hint && <span className="text-[11px]" style={{ color: colors.inkFaint }}>{hint}</span>}
    </label>
  );
}

function PackagesEditor({ form, setForm, colors, t }) {
  const [bulkText, setBulkText] = useState("");
  const [genCount, setGenCount] = useState("");
  const inputStyle = inputStyleFor(colors);
  const packages = form.packages || [];

  function addCodes() {
    const codes = bulkText.split(/[,\n]/).map((s) => s.trim()).filter(Boolean);
    if (codes.length === 0) return;
    const existing = new Set(packages.map((p) => p.code));
    const additions = codes.filter((c) => !existing.has(c)).map((c) => ({ code: c, description: "", weightKg: "", cbm: "" }));
    setForm((f) => ({ ...f, packages: [...(f.packages || []), ...additions] }));
    setBulkText("");
  }

  function generateRows() {
    const n = parseInt(genCount, 10);
    if (!n || n < 1) return;
    const startAt = packages.length;
    const additions = Array.from({ length: n }, (_, i) => ({ code: String(startAt + i + 1), description: "", weightKg: "", cbm: "" }));
    setForm((f) => ({ ...f, packages: [...(f.packages || []), ...additions] }));
    setGenCount("");
  }

  function updatePackage(idx, key, value) {
    setForm((f) => {
      const next = [...(f.packages || [])];
      next[idx] = { ...next[idx], [key]: value };
      return { ...f, packages: next };
    });
  }

  function removePackage(idx) {
    setForm((f) => ({ ...f, packages: (f.packages || []).filter((_, i) => i !== idx) }));
  }

  const totals = packages.reduce((acc, p) => ({
    weight: acc.weight + (Number(p.weightKg) || 0),
    cbm: acc.cbm + (Number(p.cbm) || 0),
  }), { weight: 0, cbm: 0 });

  return (
    <div className="mt-2">
      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionPackages}
      </div>
      <p className="text-xs mb-3" style={{ color: colors.inkFaint }}>{t.packagesHint}</p>

      <div className="flex flex-col sm:flex-row gap-2 mb-3 items-start sm:items-end">
        <Field label={t.packagesGenerateLabel} colors={colors}>
          <input
            type="number" min="1"
            className={inputClass}
            style={{ ...inputStyle, width: 90 }}
            value={genCount}
            onChange={(e) => setGenCount(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); generateRows(); } }}
          />
        </Field>
        <button
          className="px-4 py-2 rounded text-sm font-semibold h-fit"
          style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={generateRows}
        >
          {t.packagesGenerateBtn}
        </button>
      </div>

      <div className="flex flex-col md:flex-row gap-2 mb-3">
        <textarea
          className={inputClass + " flex-1"}
          style={inputStyle}
          rows={2}
          placeholder={t.bulkAddPlaceholder}
          value={bulkText}
          onChange={(e) => setBulkText(e.target.value)}
        />
        <button
          className="px-4 py-2 rounded text-sm font-semibold h-fit"
          style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={addCodes}
        >
          {t.bulkAddBtn}
        </button>
      </div>

      {packages.length === 0 ? (
        <p className="text-xs" style={{ color: colors.inkFaint }}>{t.noPackagesMsg}</p>
      ) : (
        <div className="rounded overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
          <table className="w-full text-xs" style={{ background: colors.surface }}>
            <thead>
              <tr style={{ background: colors.surfaceDim }}>
                {[t.packageCodeCol, t.packageDescCol, t.packageWeightCol, t.packageCbmCol, ""].map((h) => (
                  <th key={h} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {packages.map((p, idx) => (
                <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-1 py-1">
                    <input className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.code} onChange={(e) => updatePackage(idx, "code", e.target.value)} />
                  </td>
                  <td className="px-1 py-1">
                    <input className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.description} onChange={(e) => updatePackage(idx, "description", e.target.value)} />
                  </td>
                  <td className="px-1 py-1">
                    <input type="number" className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.weightKg} onChange={(e) => updatePackage(idx, "weightKg", e.target.value)} />
                  </td>
                  <td className="px-1 py-1">
                    <input type="number" className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.cbm || ""} onChange={(e) => updatePackage(idx, "cbm", e.target.value)} />
                  </td>
                  <td className="px-1 py-1 text-right">
                    <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => removePackage(idx)}>{t.removePackageBtn}</button>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ borderTop: `2px solid ${colors.line}`, background: colors.surfaceDim }}>
                <td className="px-2 py-1.5 font-semibold" style={{ color: colors.ink }} colSpan={2}>
                  {t.packagesTotalLabel(packages.length)}
                </td>
                <td className="px-2 py-1.5 font-semibold" style={{ color: colors.ink }}>{Math.round(totals.weight * 100) / 100}</td>
                <td className="px-2 py-1.5 font-semibold" style={{ color: colors.ink }}>{Math.round(totals.cbm * 1000) / 1000}</td>
                <td></td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </div>
  );
}

function ItemForm({ initial, onSave, onCancel, onPrintJobSheet, onMoveCases, directory, employees, currentUser, items, colors, t, lang }) {
  const [showOlderSites, setShowOlderSites] = useState(false);
  const [form, setForm] = useState(initial || { ...emptyForm(), recordedBy: currentUser || "" });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const inputStyle = inputStyleFor(colors);
  // Saving used to fail silently: the button tested these conditions and simply returned,
  // so an entry missing one just would not save and gave no reason for it. Every entry
  // created by a Devan/CFS check-in is born with an empty Description, which made all of
  // them unsaveable the moment anyone opened one to correct something.
  // A description is what tells you what an entry holds, so it is still required - but an
  // entry with an itemised case list already says that, case by case, and doesn't need one.
  const saveBlockers = [];
  if (!form.project) saveBlockers.push(t.fProject);
  if (!form.description && !(form.packages || []).length) saveBlockers.push(t.fDescription);
  if (form.depotArrivalDate && !form.recordedBy) saveBlockers.push(t.fRecordedBy);
  // The arrival date is required, because the storage ledger orders by it and storage is
  // charged from it. A blank one is not always a mistake though - an entry booked in before
  // the goods reach the depot legitimately has none, and that is what "pending collection"
  // means. So it may be left empty, but only by saying so, which turns a silent omission
  // into a deliberate state.
  if (!form.depotArrivalDate && !form.awaitingCollection) saveBlockers.push(t.fDepotArrival);
  const siteSuggestions = useMemo(() => {
    const fromDirectory = (directory || []).map((s) => s.siteEn).filter(Boolean);
    const fromItems = (items || []).map((i) => i.project).filter(Boolean);
    return [...new Set([...fromDirectory, ...fromItems])];
  }, [directory, items]);

  function applyDirectoryEntry(id) {
    const site = (directory || []).find((s) => s.id === id);
    if (!site) return;
    setForm((f) => ({
      ...f,
      directoryId: site.id,
      client: CLIENTS.includes(site.client) ? site.client : f.client,
      project: site.siteEn || f.project,
      constructionSite: site.siteZh || site.siteEn || f.constructionSite,
      jobRef: site.jobRef || f.jobRef,
      orderedBy: site.orderedBy || f.orderedBy,
    }));
  }

  return (
    <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <h3 className="text-lg font-bold mb-4" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
        {initial ? t.titleEdit : t.titleNew}
      </h3>

      {(directory || []).length > 0 && (
        <div className="mb-4">
          <Field label={t.selectFromDirectory} colors={colors}>
            <select className={inputClass} style={inputStyle} value="" onChange={(e) => applyDirectoryEntry(e.target.value)}>
              <option value="">{t.selectFromDirectoryPlaceholder}</option>
              {visibleDirectory(directory, { client: form.client, showOlder: showOlderSites, items }).map((s) => (
                <option key={s.id} value={s.id}>{s.siteEn} — {s.client}</option>
              ))}
            </select>
          </Field>
          <label className="flex items-center gap-1.5 text-xs mt-1.5" style={{ color: colors.inkFaint }}>
            <input type="checkbox" checked={showOlderSites} onChange={(e) => setShowOlderSites(e.target.checked)} />
            {t.showOlderJobs}
            {!showOlderSites && hiddenSiteCount(directory, { client: form.client, items }) > 0 && (
              <span>{t.showOlderJobsCount(hiddenSiteCount(directory, { client: form.client, items }))}</span>
            )}
          </label>
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fClient} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.client} onChange={set("client")}>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.fProjectEn} colors={colors}>
          <input list="itemform-site-suggestions" className={inputClass} style={inputStyle} placeholder={t.fProjectPlaceholder} value={form.project} onChange={set("project")} />
          <datalist id="itemform-site-suggestions">
            {siteSuggestions.map((s) => <option key={s} value={s} />)}
          </datalist>
        </Field>
        <Field label={t.fProjectZh} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.constructionSite} onChange={set("constructionSite")} />
        </Field>
        <Field label={t.fInvoiceNo} hint={t.fInvoiceHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.invoiceNumber} onChange={set("invoiceNumber")} />
        </Field>
      </div>

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionCargo}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fItemType} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.itemType} onChange={set("itemType")}>
            {ITEM_TYPES.map((tt) => <option key={tt}>{tt}</option>)}
          </select>
        </Field>
        {form.itemType === "Separate Items" && (!form.packages || form.packages.length === 0) && (
          <Field label={t.fPackages} colors={colors}>
            <input type="number" min="0" className={inputClass} style={inputStyle} value={form.packageCount} onChange={set("packageCount")} />
          </Field>
        )}
        {form.itemType === "Separate Items" && form.packages && form.packages.length > 0 && (
          <Field label={t.fPackages} colors={colors}>
            <div className="px-2.5 py-1.5 rounded text-sm" style={{ ...inputStyle, background: colors.surfaceDim }}>
              {t.packagesCountSummary(form.packages.length)}
            </div>
          </Field>
        )}
        <Field label={t.fUnitCode} hint={t.fUnitCodeHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.unitCode} onChange={set("unitCode")} />
        </Field>

        <Field label={t.fDescription} hint={t.fDescriptionHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.description} onChange={set("description")} />
        </Field>
        <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.shkNumber} onChange={set("shkNumber")} />
        </Field>
        <Field label={t.fSsDoNo} hint={t.fSsDoNoHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.ssDoNo || ""} onChange={set("ssDoNo")} placeholder={'ex ss."SHIP" V.___; CONTAINERS NO. ___'} />
        </Field>

        <Field
          label={t.fWeight}
          hint={form.weightSource
            ? `${form.weightSource === "declared-estimated" ? t.fWeightSheetShare : t.fWeightFromSheet}${form.weightPackingListKg ? ` \u00b7 ${t.fWeightPackingList(form.weightPackingListKg)}` : ""}`
            : undefined}
          colors={colors}
        >
          <input type="number" className={inputClass} style={inputStyle} value={form.weightKg} onChange={set("weightKg")} />
        </Field>
        <Field
          label={t.fVolume}
          hint={form.volumeSource
            ? `${form.volumeSource === "oversize" ? t.fVolumeFromOversize : form.volumeSource === "declared-estimated" ? t.fWeightSheetShare : t.fWeightFromSheet}${form.volumeCbmPackingList ? ` \u00b7 ${t.fVolumePackingList(form.volumeCbmPackingList)}` : ""}`
            : undefined}
          colors={colors}
        >
          <input type="number" className={inputClass} style={inputStyle} value={form.volumeCbm} onChange={set("volumeCbm")} />
        </Field>
        <div />

        <Field label={t.f20} colors={colors}>
          <input type="number" min="0" className={inputClass} style={inputStyle} value={form.containers20} onChange={set("containers20")} />
        </Field>
        <Field label={t.f40} colors={colors}>
          <input type="number" min="0" className={inputClass} style={inputStyle} value={form.containers40} onChange={set("containers40")} />
        </Field>
        <div />

        <Field label={t.fZone} hint={t.fZoneHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.zone || "urban"} onChange={set("zone")}>
            <option value="urban">{t.zoneUrban}</option>
            <option value="lantau">{t.zoneLantau}</option>
          </select>
        </Field>
        <Field label={t.fCargoType} hint={t.fCargoTypeHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.cargoType || ""} onChange={set("cargoType")}>
            <option value="">{t.cargoTypeAuto}</option>
            <option value="elevator">{t.cargoTypeElevator}</option>
            <option value="escalator">{t.cargoTypeEscalator}</option>
          </select>
        </Field>
        <div />
      </div>

      {form.itemType === "Separate Items" && (
        <PackagesEditor form={form} setForm={setForm} colors={colors} t={t} />
      )}

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionArrival}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fArrivingType} hint={t.fArrivingTypeHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.arrivingType} onChange={set("arrivingType")}>
            {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
          </select>
        </Field>
        <Field label={t.fDepot} hint={t.fDepotHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.depot} onChange={set("depot")}>
            {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
          </select>
        </Field>
        <Field label={t.fDepotLocation} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.depotLocation} onChange={set("depotLocation")} />
        </Field>

        <Field label={t.fTerminalArrival} hint={t.fTerminalArrivalHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.terminalArrivalDate} onChange={set("terminalArrivalDate")} />
        </Field>
        <Field label={t.fTerminalLFD} hint={t.fTerminalLFDHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.terminalLFD} onChange={set("terminalLFD")} />
        </Field>
        <Field label={t.fConfirmedCollection} hint={t.fConfirmedCollectionHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.confirmedCollectionDate} onChange={set("confirmedCollectionDate")} />
        </Field>

        <Field label={t.fDepotArrival} hint={t.fDepotArrivalHint} colors={colors}>
          <input type="date" className={inputClass}
            style={{ ...inputStyle, ...(form.depotArrivalDate || form.awaitingCollection ? {} : { borderColor: colors.red, background: colors.redSoft }) }}
            value={form.depotArrivalDate} onChange={set("depotArrivalDate")} />
        </Field>
        <Field label={t.fPlannedDelivery} hint={t.fPlannedDeliveryHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.plannedDeliveryDate} onChange={set("plannedDeliveryDate")} />
        </Field>
        <div />
        <div />

        {(form.packages || []).length > 0 && (
          <ArrivalBatchesEditor form={form} setForm={setForm} colors={colors} t={t} lang={lang} />
        )}

        {onMoveCases && form.id && (form.packages || []).length > 0 && (
          <MoveCasesEditor
            form={form} setForm={setForm} items={items} onMoveCases={onMoveCases}
            colors={colors} t={t} inputClass={inputClass}
          />
        )}

        {(form.deliveries || []).length > 0 && (
          <div className="col-span-2 md:col-span-3 px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>
            {t.deliveryProgress(deliveredUnits(form), totalUnits(form), form.deliveries.length)}
            <strong>{fmt(lastDeliveryDate(form))}</strong>. {t.deliveryProgressManage}
          </div>
        )}

        <div className="col-span-2 md:col-span-3">
          <Field label={t.fNotes} colors={colors}>
            <textarea className={inputClass} style={inputStyle} rows={2} value={form.notes} onChange={set("notes")} />
          </Field>
        </div>
      </div>

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionJobSheet}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
          <div className="flex gap-2">
            <input className={inputClass + " flex-1"} style={inputStyle} value={form.jobNumber} onChange={set("jobNumber")} />
            <button
              type="button"
              className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => setForm((f) => ({ ...f, jobNumber: nextJobNumber(items) }))}
            >
              {t.generateJobNoBtn}
            </button>
          </div>
        </Field>
        <Field label={t.fOrderedBy} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.orderedBy} onChange={set("orderedBy")} />
        </Field>
        <Field label={t.fPoNumber} hint={t.fPoNumberHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.poNumber} onChange={set("poNumber")} />
        </Field>
        <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.jobRef} onChange={set("jobRef")} />
        </Field>
        <Field label={t.fRecordedBy} hint={t.fRecordedByHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.recordedBy} onChange={set("recordedBy")}>
            <option value=""></option>
            {(employees || []).map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
          </select>
        </Field>
        {form.jobNumber && (
          <div className="col-span-2 md:col-span-3">
            <button
              type="button"
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ border: `1px solid ${colors.line}`, color: colors.amberText, fontFamily: FONT_DISPLAY }}
              onClick={() => onPrintJobSheet({ type: form.arrivingType, item: form })}
            >
              {t.printJobSheetBtn}
            </button>
          </div>
        )}
      </div>

      {!form.depotArrivalDate && (
        <label className="flex items-center gap-2 text-xs cursor-pointer" style={{ color: colors.amberText }}>
          <input type="checkbox" checked={!!form.awaitingCollection}
            onChange={(e) => setForm((f) => ({ ...f, awaitingCollection: e.target.checked }))} />
          {t.fAwaitingCollection}
        </label>
      )}
      {form.depotArrivalDate && !form.recordedBy && (
        <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
          {t.recordedByRequiredMsg}
        </div>
      )}

      {saveBlockers.filter((b) => b !== t.fRecordedBy).length > 0 && (
        <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
          {t.saveBlockedMsg(saveBlockers.filter((b) => b !== t.fRecordedBy).join(", "))}
        </div>
      )}

      <div className="flex gap-2 mt-5">
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{
            background: saveBlockers.length ? colors.line : colors.navy,
            color: saveBlockers.length ? colors.inkFaint : colors.onDark,
            fontFamily: FONT_DISPLAY,
            cursor: saveBlockers.length ? "not-allowed" : "pointer",
          }}
          onClick={() => {
            if (saveBlockers.length) return;
            onSave(form);
          }}
        >
          {t.saveBtn}
        </button>
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={onCancel}
        >
          {t.cancelBtn}
        </button>
      </div>
    </div>
  );
}

// Moves cases out of this entry and into another. A lot filed under the wrong lift used to
// mean deleting the entry and rebuilding it from the packing list forward; this repairs it
// in place. Only cases still at the depot are offered - one that has already gone out
// belongs to a delivery record, and pulling it out of that record would rewrite history.
function MoveCasesEditor({ form, setForm, items, onMoveCases, colors, t, inputClass }) {
  const [picked, setPicked] = useState([]);
  const [destId, setDestId] = useState("");
  const [done, setDone] = useState("");
  const inputStyle = inputStyleFor(colors);
  const movable = remainingPackages(form);
  // Somewhere sensible to move to: the same client, and either the same job number or the
  // same site - the sibling lots that came in on one packing list.
  const destinations = (items || []).filter((i) =>
    i.id !== form.id && !i.cancelled && i.client === form.client &&
    ((form.jobNumber && String(i.jobNumber || "").trim() === String(form.jobNumber).trim())
      || sitesLooselyMatch(form.project, form.constructionSite, i.project, i.constructionSite)));
  if (!movable.length || !destinations.length) return null;

  function move() {
    const dest = destinations.find((d) => d.id === destId);
    if (!dest || !picked.length) return;
    onMoveCases({ fromId: form.id, toId: dest.id, codes: picked });
    // The form holds its own copy of the entry, so it has to lose the cases too or it
    // would put them back the next time it is saved.
    const gone = new Set(picked);
    setForm((f) => ({
      ...f,
      packages: (f.packages || []).filter((p) => !gone.has(p.code)),
      arrivals: (f.arrivals || []).map((a) => ({ ...a, codes: (a.codes || []).filter((c) => !gone.has(c)) })),
    }));
    setDone(t.moveCasesDoneMsg(picked.join(", "), dest.unitCode || dest.id));
    setPicked([]);
  }

  return (
    <div className="col-span-2 md:col-span-3 rounded p-3" style={{ border: `1px dashed ${colors.line}` }}>
      <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
        {t.moveCasesLabel}
      </div>
      <div className="text-[11px] mb-2" style={{ color: colors.inkFaint }}>{t.moveCasesHint}</div>
      <div className="flex flex-wrap gap-2 mb-2">
        {movable.map((p) => {
          const on = picked.includes(p.code);
          return (
            <button
              key={p.code}
              type="button"
              onClick={() => setPicked((prev) => on ? prev.filter((c) => c !== p.code) : [...prev, p.code])}
              className="px-2.5 py-1.5 rounded text-xs font-semibold"
              style={{
                border: `1px solid ${on ? colors.amber : colors.line}`,
                background: on ? colors.amberSoft : colors.surface,
                color: on ? colors.amberText : colors.ink,
              }}
              title={p.description}
            >
              {p.code}
            </button>
          );
        })}
      </div>
      <div className="flex flex-wrap items-center gap-2">
        <select className={inputClass} style={{ ...inputStyle, width: 260, fontSize: 12, padding: "4px 8px" }}
          value={destId} onChange={(e) => { setDestId(e.target.value); setDone(""); }}>
          <option value="">{t.moveCasesDestPlaceholder}</option>
          {destinations.map((d) => (
            <option key={d.id} value={d.id}>{`${d.id}${d.unitCode ? ` \u00b7 ${d.unitCode}` : ""}`}</option>
          ))}
        </select>
        <button
          type="button"
          className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{
            background: picked.length && destId ? colors.amber : colors.line,
            color: picked.length && destId ? colors.ink : colors.inkFaint,
            fontFamily: FONT_DISPLAY,
            cursor: picked.length && destId ? "pointer" : "not-allowed",
          }}
          onClick={move}
        >
          {t.moveCasesBtn}
        </button>
      </div>
      {done && <div className="text-xs mt-2" style={{ color: colors.green }}>{done}</div>}
    </div>
  );
}
function ArrivalBatchesEditor({ form, setForm, colors, t, lang }) {
  const inputStyle = inputStyleFor(colors);
  const [draft, setDraft] = useState({ date: todayStr(), type: ARRIVING_TYPES[0], codes: [] });
  const arrivals = form.arrivals || [];
  const assigned = new Set(arrivals.flatMap((a) => a.codes || []));
  const unassigned = (form.packages || []).filter((p) => !assigned.has(p.code));

  function toggleDraftCode(code) {
    setDraft((d) => ({ ...d, codes: d.codes.includes(code) ? d.codes.filter((c) => c !== code) : [...d.codes, code] }));
  }
  function syncEarliestDate(list) {
    const ds = list.map((a) => a.date).filter(Boolean).sort();
    return ds.length ? ds[0] : "";
  }
  function addBatch() {
    if (!draft.date || draft.codes.length === 0) return;
    const next = [...arrivals, { id: `arr-${Date.now()}`, date: draft.date, type: draft.type, codes: draft.codes }];
    setForm((f) => ({ ...f, arrivals: next, depotArrivalDate: syncEarliestDate(next) }));
    setDraft({ date: todayStr(), type: ARRIVING_TYPES[0], codes: [] });
  }
  function removeBatch(id) {
    const next = arrivals.filter((a) => a.id !== id);
    setForm((f) => ({ ...f, arrivals: next, depotArrivalDate: next.length ? syncEarliestDate(next) : f.depotArrivalDate }));
  }
  // A batch recorded from a Devan/CFS job sheet was previously fixed for good: a mistyped
  // date or a weight read off the wrong line meant deleting the entry and re-uploading
  // everything. Each batch is editable here instead.
  function patchBatch(id, patch) {
    const next = arrivals.map((a) => (a.id === id ? { ...a, ...patch } : a));
    setForm((f) => ({ ...f, arrivals: next, depotArrivalDate: syncEarliestDate(next) }));
  }
  // A figure typed in by hand is a figure stated for this lot, so it drops the "split"
  // flag it may have carried. That flag marks a share of a total covering several lots,
  // which is treated as an estimate and loses to per-case packing-list weights - leaving
  // it set would mean the number someone just corrected still wasn't the one used.
  function patchBatchDeclared(id, key, value) {
    const cur = arrivals.find((a) => a.id === id);
    const declared = { ...(cur && cur.declared ? cur.declared : {}), [key]: value, split: false };
    const any = ["pkgs", "kg", "cbm"].some((k) => String(declared[k] || "").trim() !== "");
    patchBatch(id, { declared: any ? declared : null, declaredEdited: true });
  }

  return (
    <div className="col-span-2 md:col-span-3 rounded p-3" style={{ border: `1px dashed ${colors.line}` }}>
      <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
        {t.splitArrivalLabel}
      </div>
      <div className="text-[11px] mb-2" style={{ color: colors.inkFaint }}>{t.splitArrivalHint}</div>

      {arrivals.length > 0 && (
        <div className="mb-3 rounded overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
          <table className="w-full text-xs" style={{ background: colors.surface }}>
            <thead>
              <tr style={{ background: colors.surfaceDim }}>
                {[t.colDate, t.fArrivingType, t.splitArrivalCasesCol, t.jsKgs, t.jsCbm, ""].map((h, idx) => (
                  <th key={idx} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...arrivals].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((a) => (
                <tr key={a.id} style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-2 py-1.5">
                    <input type="date" className={inputClass} style={{ ...inputStyle, fontSize: 12, padding: "3px 6px" }}
                      value={a.date || ""} onChange={(e) => patchBatch(a.id, { date: e.target.value })} />
                  </td>
                  <td className="px-2 py-1.5">
                    <select className={inputClass} style={{ ...inputStyle, fontSize: 12, padding: "3px 6px" }}
                      value={a.type || ARRIVING_TYPES[0]} onChange={(e) => patchBatch(a.id, { type: e.target.value })}>
                      {ARRIVING_TYPES.map((x) => <option key={x}>{x}</option>)}
                    </select>
                  </td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{(a.codes || []).join(", ")}</td>
                  <td className="px-2 py-1.5">
                    <input type="number" min="0" step="0.1" className={inputClass}
                      style={{ ...inputStyle, width: 88, fontSize: 12, padding: "3px 6px" }}
                      value={(a.declared && a.declared.kg) || ""}
                      onChange={(e) => patchBatchDeclared(a.id, "kg", e.target.value)} />
                  </td>
                  <td className="px-2 py-1.5">
                    <input type="number" min="0" step="0.001" className={inputClass}
                      style={{ ...inputStyle, width: 84, fontSize: 12, padding: "3px 6px" }}
                      value={(a.declared && a.declared.cbm) || ""}
                      onChange={(e) => patchBatchDeclared(a.id, "cbm", e.target.value)} />
                  </td>
                  <td className="px-2 py-1.5 text-right">
                    <button type="button" className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => removeBatch(a.id)}>{t.deleteBtn}</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {unassigned.length > 0 ? (
        <div className="flex flex-col gap-2">
          <div className="flex flex-wrap items-end gap-3">
            <Field label={t.colDate} colors={colors}>
              <input type="date" className={inputClass} style={inputStyle} value={draft.date} onChange={(e) => setDraft((d) => ({ ...d, date: e.target.value }))} />
            </Field>
            <Field label={t.fArrivingType} colors={colors}>
              <select className={inputClass} style={inputStyle} value={draft.type} onChange={(e) => setDraft((d) => ({ ...d, type: e.target.value }))}>
                {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
              </select>
            </Field>
            <button
              type="button"
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: draft.codes.length === 0 ? 0.5 : 1 }}
              disabled={draft.codes.length === 0}
              onClick={addBatch}
            >
              {t.splitArrivalAddBtn}
            </button>
          </div>
          <div className="text-[11px]" style={{ color: colors.inkFaint }}>{t.splitArrivalSelectHint}</div>
          <div className="flex flex-wrap gap-2">
            {unassigned.map((p) => (
              <button
                key={p.code}
                type="button"
                onClick={() => toggleDraftCode(p.code)}
                className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                style={{
                  border: `1px solid ${draft.codes.includes(p.code) ? colors.amber : colors.line}`,
                  background: draft.codes.includes(p.code) ? colors.amberSoft : colors.surface,
                  color: draft.codes.includes(p.code) ? colors.amberText : colors.ink,
                }}
                title={p.description}
              >
                {p.code}
              </button>
            ))}
          </div>
        </div>
      ) : arrivals.length > 0 ? (
        <div className="text-xs" style={{ color: colors.green }}>{t.splitArrivalAllAssignedMsg}</div>
      ) : null}
    </div>
  );
}

// Corrects a delivery that has already been recorded. The cases it may hold are the ones
// it took plus whatever is still at the depot on this entry - anything else belongs to a
// different delivery and is not this record's to claim.
function DeliveryRecordEditor({ delivery, item, onPatch, onDone, colors, t }) {
  const inputStyle = inputStyleFor(colors);
  const [draft, setDraft] = useState({
    date: delivery.date || "",
    deliveredTo: delivery.deliveredTo || "",
    receivedBy: delivery.receivedBy || "",
    jobNumber: delivery.jobNumber || "",
    kg: (delivery.declared && delivery.declared.kg) || "",
    cbm: (delivery.declared && delivery.declared.cbm) || "",
    codes: [...(delivery.codes || [])],
  });
  const mine = new Set(delivery.codes || []);
  const choosable = [
    ...(item.packages || []).filter((p) => mine.has(p.code)),
    ...remainingPackages(item).filter((p) => !mine.has(p.code)),
  ];
  const set = (k) => (e) => setDraft((d) => ({ ...d, [k]: e.target.value }));
  function save() {
    const patch = {
      date: draft.date, deliveredTo: draft.deliveredTo, receivedBy: draft.receivedBy,
      jobNumber: draft.jobNumber,
    };
    if (delivery.codes) patch.codes = draft.codes;
    // A figure corrected by hand is stated for this delivery, not a share of a bigger
    // total, so it drops the split flag that would otherwise mark it an estimate.
    const kg = String(draft.kg).trim(), cbm = String(draft.cbm).trim();
    patch.declared = (kg || cbm) ? { kg, cbm, split: false } : null;
    onPatch(patch);
    onDone();
  }
  return (
    <div className="flex flex-col gap-2">
      <div className="flex flex-wrap items-end gap-2">
        <Field label={t.colDate} colors={colors}>
          <input type="date" className={inputClass} style={{ ...inputStyle, fontSize: 12, padding: "3px 6px" }} value={draft.date} onChange={set("date")} />
        </Field>
        <Field label={t.colDeliveredTo} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 200, fontSize: 12, padding: "3px 6px" }} value={draft.deliveredTo} onChange={set("deliveredTo")} />
        </Field>
        <Field label={t.colReceivedBy} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 150, fontSize: 12, padding: "3px 6px" }} value={draft.receivedBy} onChange={set("receivedBy")} />
        </Field>
        <Field label={t.colJobNo} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 110, fontSize: 12, padding: "3px 6px" }} value={draft.jobNumber} onChange={set("jobNumber")} />
        </Field>
        <Field label={t.jsKgs} colors={colors}>
          <input type="number" min="0" step="0.1" className={inputClass} style={{ ...inputStyle, width: 90, fontSize: 12, padding: "3px 6px" }} value={draft.kg} onChange={set("kg")} />
        </Field>
        <Field label={t.jsCbm} colors={colors}>
          <input type="number" min="0" step="0.001" className={inputClass} style={{ ...inputStyle, width: 86, fontSize: 12, padding: "3px 6px" }} value={draft.cbm} onChange={set("cbm")} />
        </Field>
      </div>
      {delivery.codes && choosable.length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {choosable.map((p) => {
            const on = draft.codes.includes(p.code);
            return (
              <button
                key={p.code}
                type="button"
                onClick={() => setDraft((d) => ({ ...d, codes: on ? d.codes.filter((c) => c !== p.code) : [...d.codes, p.code] }))}
                className="px-2 py-1 rounded text-xs font-semibold"
                style={{
                  border: `1px solid ${on ? colors.amber : colors.line}`,
                  background: on ? colors.amberSoft : colors.surface,
                  color: on ? colors.amberText : colors.ink,
                }}
                title={p.description}
              >
                {p.code}
              </button>
            );
          })}
        </div>
      )}
      <div className="flex gap-2">
        <button type="button" className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
          onClick={save}>{t.saveBtn}</button>
        <button type="button" className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={onDone}>{t.cancelBtn}</button>
      </div>
    </div>
  );
}
function DeliveryForm({ deliveryItems, onAddDelivery, onAddCombinedDelivery, onDeleteDelivery, onUpdateDelivery, onCancel, onPrintJobSheet, employees, currentUser, items, colors, t, lang }) {
  const firstItem = deliveryItems[0];
  const [extraItemIds, setExtraItemIds] = useState([]);
  const [editingDeliveryId, setEditingDeliveryId] = useState(null);
  const [addPickerOpen, setAddPickerOpen] = useState(false);
  const [addSearch, setAddSearch] = useState("");
  const baseIds = new Set(deliveryItems.map((i) => i.id));
  const allDeliveryItems = [...deliveryItems, ...items.filter((i) => extraItemIds.includes(i.id) && !baseIds.has(i.id))];
  const isCombined = allDeliveryItems.length > 1;
  const addCandidates = useMemo(() => {
    const excluded = new Set(allDeliveryItems.map((i) => i.id));
    const siteKey = (i) => (i.constructionSite || i.project || "").trim().toLowerCase();
    const targetSite = siteKey(firstItem);
    return items.filter((i) => !excluded.has(i.id) && ["at_depot", "partial"].includes(deriveStatus(i)) && (!targetSite || siteKey(i) === targetSite));
  }, [items, allDeliveryItems, firstItem]);
  const addFiltered = addSearch.trim()
    ? addCandidates.filter((i) => {
        const q = addSearch.toLowerCase();
        return i.unitCode?.toLowerCase().includes(q) || i.jobNumber?.toLowerCase().includes(q) || i.id?.toLowerCase().includes(q);
      })
    : addCandidates;

  const [form, setForm] = useState({
    date: todayStr(),
    deliveredTo: firstItem.constructionSite || firstItem.project || "",
    receivedBy: "",
    notes: "",
    jobNumber: "",
    recordedBy: currentUser || "",
  });
  const [perItem, setPerItem] = useState({});
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const inputStyle = inputStyleFor(colors);

  function getSel(itemId) {
    return perItem[itemId] || { codes: [], qty: "" };
  }
  function setCodes(itemId, updater) {
    setPerItem((prev) => {
      const cur = prev[itemId] || { codes: [], qty: "" };
      const nextCodes = typeof updater === "function" ? updater(cur.codes) : updater;
      return { ...prev, [itemId]: { ...cur, codes: nextCodes } };
    });
  }
  function setQty(itemId, qty) {
    setPerItem((prev) => ({ ...prev, [itemId]: { ...(prev[itemId] || { codes: [] }), qty } }));
  }
  function toggleCode(itemId, code) {
    setCodes(itemId, (prev) => (prev.includes(code) ? prev.filter((c) => c !== code) : [...prev, code]));
  }
  function toggleBatch(itemId, batch) {
    setCodes(itemId, (prev) => {
      const allSelected = batch.remainingCodes.every((c) => prev.includes(c));
      if (allSelected) return prev.filter((c) => !batch.remainingCodes.includes(c));
      return [...new Set([...prev, ...batch.remainingCodes])];
    });
  }
  function addBatchToDelivery(itemId) {
    setExtraItemIds((prev) => (prev.includes(itemId) ? prev : [...prev, itemId]));
    setAddSearch("");
    setAddPickerOpen(false);
  }
  function removeExtraBatch(itemId) {
    setExtraItemIds((prev) => prev.filter((id) => id !== itemId));
    setPerItem((prev) => { const next = { ...prev }; delete next[itemId]; return next; });
  }

  const anySelected = allDeliveryItems.some((it) => {
    const itemized = (it.packages || []).length > 0;
    const sel = getSel(it.id);
    return itemized ? sel.codes.length > 0 : Number(sel.qty) > 0;
  });
  const anyOvershoot = allDeliveryItems.some((it) => {
    const itemized = (it.packages || []).length > 0;
    if (itemized) return false;
    const sel = getSel(it.id);
    return Number(sel.qty) > remainingUnits(it);
  });

  function handleAddClick() {
    const jobNumber = form.jobNumber || nextJobNumber(items);
    const deliveredTo = form.deliveredTo || firstItem.constructionSite || firstItem.project || "";
    const receivedBy = form.receivedBy || firstItem.orderedBy || "";
    const entries = [];
    for (const it of allDeliveryItems) {
      const itemized = (it.packages || []).length > 0;
      const sel = getSel(it.id);
      if (itemized) {
        if (sel.codes.length === 0) continue;
        entries.push({ itemId: it.id, delivery: { date: form.date, deliveredTo, receivedBy, jobNumber, recordedBy: form.recordedBy, notes: form.notes, codes: sel.codes } });
      } else {
        const qty = Number(sel.qty) || 0;
        if (qty <= 0) continue;
        entries.push({ itemId: it.id, delivery: { date: form.date, deliveredTo, receivedBy, jobNumber, recordedBy: form.recordedBy, notes: form.notes, packageCount: qty } });
      }
    }
    if (entries.length === 0) return;
    setPerItem({});
    setExtraItemIds([]);
    setForm({ date: todayStr(), deliveredTo: firstItem.constructionSite || firstItem.project || "", receivedBy: "", notes: "", jobNumber: "", recordedBy: currentUser || "" });
    if (entries.length === 1 && !isCombined) {
      const record = onAddDelivery(entries[0].delivery, entries[0].itemId);
      onPrintJobSheet({ type: "Delivery", item: firstItem, delivery: record || entries[0].delivery });
    } else {
      const recorded = onAddCombinedDelivery(entries);
      const groups = recorded.map(({ itemId, record }) => ({ item: allDeliveryItems.find((i) => i.id === itemId), delivery: record }));
      onPrintJobSheet({ type: "Delivery", combined: true, groups, jobNumber, date: form.date, deliveredTo, receivedBy });
    }
  }

  return (
    <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
        {isCombined ? t.combinedDeliveryTitle(allDeliveryItems.length) : `${t.deliveryTitlePrefix} ${firstItem.id}`}
      </h3>
      <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>
        {isCombined
          ? allDeliveryItems.map((it) => `${it.unitCode || it.id}`).join(" · ")
          : <>{firstItem.client} · {firstItem.project}{firstItem.plannedDeliveryDate ? t.plannedWasText(fmt(firstItem.plannedDeliveryDate)) : ""}</>}
      </p>

      {!isCombined && activeDeliveries(firstItem).length > 0 && (
        <div className="mb-4 rounded overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
          <table className="w-full text-xs" style={{ background: colors.surface }}>
            <thead>
              <tr style={{ background: colors.surfaceDim }}>
                {[t.colDate, t.colQty, t.colDeliveredTo, t.colReceivedBy, t.colJobNo, ""].map((h) => (
                  <th key={h} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...activeDeliveries(firstItem)].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((d) => (
                <React.Fragment key={d.id}>
                <tr style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{fmt(d.date)}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.codes ? d.codes.join(", ") : d.packageCount}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.deliveredTo || "—"}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.receivedBy || "—"}</td>
                  <td className="px-2 py-1.5" style={{ fontFamily: FONT_MONO, color: colors.ink }}>{d.jobNumber || "—"}</td>
                  <td className="px-2 py-1.5 text-right whitespace-nowrap">
                    <button className="text-xs font-semibold mr-2" style={{ color: colors.amberText }} onClick={() => onPrintJobSheet({ type: "Delivery", item: firstItem, delivery: d })}>{t.printBtn}</button>
                    {onUpdateDelivery && (
                      <button className="text-xs font-semibold mr-2" style={{ color: colors.amberText }}
                        onClick={() => setEditingDeliveryId(editingDeliveryId === d.id ? null : d.id)}>{t.editBtn}</button>
                    )}
                    <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => onDeleteDelivery(d.id, firstItem.id)}>{t.cancelJobBtn}</button>
                  </td>
                </tr>
                {editingDeliveryId === d.id && onUpdateDelivery && (
                  <tr style={{ background: colors.surfaceDim }}>
                    <td colSpan={6} className="px-2 py-2">
                      <DeliveryRecordEditor
                        delivery={d} item={firstItem}
                        onPatch={(patch) => onUpdateDelivery(d.id, firstItem.id, patch)}
                        onDone={() => setEditingDeliveryId(null)}
                        colors={colors} t={t}
                      />
                    </td>
                  </tr>
                )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {isCombined && (
        <div className="mb-4 px-3 py-2 rounded text-xs" style={{ background: colors.surfaceDim, color: colors.inkFaint }}>
          {t.combinedHistoryHiddenNote}
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fDeliveryDate} hint={t.fDeliveryDateHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.date} onChange={set("date")} />
        </Field>
        <div />
        <div />
        <Field label={t.fDeliveredTo} hint={t.fDeliveredToHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.deliveredTo} onChange={set("deliveredTo")} />
        </Field>
        <Field label={t.fReceivedBy} hint={t.fReceivedByHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.receivedBy} onChange={set("receivedBy")} />
        </Field>
        <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
          <div className="flex gap-2">
            <input className={inputClass + " flex-1"} style={inputStyle} value={form.jobNumber} onChange={set("jobNumber")} />
            <button
              type="button"
              className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => setForm((f) => ({ ...f, jobNumber: nextJobNumber(items) }))}
            >
              {t.generateJobNoBtn}
            </button>
          </div>
        </Field>
        <Field label={t.fRecordedBy} hint={t.fRecordedByHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.recordedBy} onChange={set("recordedBy")}>
            <option value=""></option>
            {(employees || []).map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
          </select>
        </Field>
        <div className="col-span-2 md:col-span-3">
          <Field label={t.fNotes} colors={colors}>
            <textarea className={inputClass} style={inputStyle} rows={2} value={form.notes} onChange={set("notes")} />
          </Field>
        </div>
      </div>

      {allDeliveryItems.map((it) => {
        const remaining = remainingUnits(it);
        const multiUnit = totalUnits(it) > 1;
        const itemized = (it.packages || []).length > 0;
        const remainingPkgs = remainingPackages(it);
        const pendingArrival = new Set(notYetArrivedPackages(it).map((p) => p.code));
        const remainingCodesSet = new Set(remainingPkgs.map((p) => p.code));
        const selectableBatches = usesArrivalBatches(it)
          ? activeArrivals(it).map((b) => ({ ...b, remainingCodes: (b.codes || []).filter((c) => remainingCodesSet.has(c)) })).filter((b) => b.remainingCodes.length > 0)
          : [];
        const sel = getSel(it.id);

        if (remaining <= 0) {
          return (
            <div key={it.id} className="mt-4 px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>
              {(it.unitCode || it.id)}: {itemized ? t.noCodesRemainingMsg : t.allDeliveredMsg}
            </div>
          );
        }

        return (
          <div key={it.id} className="mt-5 pt-4" style={{ borderTop: `1px solid ${colors.line}` }}>
            {isCombined && (
              <div className="text-sm font-bold mb-2" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
                {it.unitCode || it.id} · {it.client} · {it.project}{it.jobNumber ? ` · ${t.colJobNo}: ${it.jobNumber}` : ""}
              </div>
            )}
            {itemized ? (
              <div>
                {selectableBatches.length > 1 && (
                  <div className="mb-3">
                    <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                      {t.selectByBatchLabel}
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {selectableBatches.map((b) => {
                        const allSelected = b.remainingCodes.every((c) => sel.codes.includes(c));
                        const someSelected = !allSelected && b.remainingCodes.some((c) => sel.codes.includes(c));
                        return (
                          <button
                            key={b.id}
                            type="button"
                            onClick={() => toggleBatch(it.id, b)}
                            className="px-3 py-1.5 rounded text-xs font-semibold text-left"
                            style={{
                              border: `1px solid ${allSelected || someSelected ? colors.amber : colors.line}`,
                              background: allSelected ? colors.amberSoft : someSelected ? colors.surfaceDim : colors.surface,
                              color: allSelected || someSelected ? colors.amberText : colors.ink,
                            }}
                          >
                            {fmt(b.date)}{b.type ? ` · ${b.type}` : ""} · {b.remainingCodes.length} {t.jsPkgs}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
                <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                  {t.selectCodesLabel}
                </div>
                <div className="flex flex-wrap gap-2">
                  {remainingPkgs.map((p) => {
                    const notHere = pendingArrival.has(p.code);
                    if (notHere) {
                      return (
                        <span
                          key={p.code}
                          className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                          style={{ border: `1px dashed ${colors.line}`, background: colors.surfaceDim, color: colors.inkFaint, cursor: "not-allowed" }}
                          title={t.notYetArrivedHint}
                        >
                          {p.code} · {t.notYetArrivedTag}
                        </span>
                      );
                    }
                    return (
                    <button
                      key={p.code}
                      type="button"
                      onClick={() => toggleCode(it.id, p.code)}
                      className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                      style={{
                        border: `1px solid ${sel.codes.includes(p.code) ? colors.amber : colors.line}`,
                        background: sel.codes.includes(p.code) ? colors.amberSoft : colors.surface,
                        color: sel.codes.includes(p.code) ? colors.amberText : colors.ink,
                      }}
                      title={p.description}
                    >
                      {p.code}{p.description ? ` — ${p.description}` : ""}
                    </button>
                    );
                  })}
                </div>
                {pendingArrival.size > 0 && (
                  <div className="mt-2 px-3 py-2 rounded text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
                    {t.pendingArrivalNotice(pendingArrival.size)}
                  </div>
                )}
              </div>
            ) : (
              <div className="max-w-xs">
                {multiUnit ? (
                  <Field label={t.fQty} hint={t.fQtyHint(remaining)} colors={colors}>
                    <input type="number" min="1" max={remaining} className={inputClass} style={inputStyle} value={sel.qty} onChange={(e) => setQty(it.id, e.target.value)} />
                  </Field>
                ) : (
                  <button
                    type="button"
                    className="px-2.5 py-1.5 rounded text-xs font-semibold"
                    style={{
                      border: `1px solid ${Number(sel.qty) > 0 ? colors.amber : colors.line}`,
                      background: Number(sel.qty) > 0 ? colors.amberSoft : colors.surface,
                      color: Number(sel.qty) > 0 ? colors.amberText : colors.ink,
                    }}
                    onClick={() => setQty(it.id, Number(sel.qty) > 0 ? "" : "1")}
                  >
                    {Number(sel.qty) > 0 ? t.selectedTag : t.includeInDeliveryBtn}
                  </button>
                )}
                {Number(sel.qty) > remaining && (
                  <div className="mt-2 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
                    {t.overshootMsg(remaining)}
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}

      <div className="mt-5 pt-4" style={{ borderTop: `1px solid ${colors.line}` }}>
        {extraItemIds.length > 0 && (
          <div className="flex flex-wrap gap-2 mb-3">
            {items.filter((i) => extraItemIds.includes(i.id)).map((i) => (
              <span key={i.id} className="inline-flex items-center gap-2 px-2.5 py-1.5 rounded text-xs font-semibold" style={{ background: colors.amberSoft, color: colors.amberText, border: `1px solid ${colors.amber}` }}>
                {i.unitCode || i.id} · {i.client}
                <button type="button" onClick={() => removeExtraBatch(i.id)} style={{ color: colors.amberText, fontWeight: "bold" }}>&times;</button>
              </span>
            ))}
          </div>
        )}
        {!addPickerOpen ? (
          <button
            type="button"
            className="px-3 py-2 rounded text-sm font-semibold"
            style={{ border: `1px dashed ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
            onClick={() => setAddPickerOpen(true)}
          >
            {t.addMoreBatchesBtn}
          </button>
        ) : (
          <div className="rounded p-3" style={{ border: `1px solid ${colors.line}`, background: colors.surfaceDim }}>
            <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>
              {t.addMoreBatchesScopeNote(firstItem.constructionSite || firstItem.project || "")}
            </div>
            <div className="flex items-center gap-2 mb-2">
              <input
                autoFocus
                className={inputClass}
                style={{ ...inputStyle, flex: 1 }}
                placeholder={t.addMoreBatchesSearchPlaceholder}
                value={addSearch}
                onChange={(e) => setAddSearch(e.target.value)}
              />
              <button type="button" className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => { setAddPickerOpen(false); setAddSearch(""); }}>{t.closeBtn}</button>
            </div>
            <div className="flex flex-col gap-1 max-h-56 overflow-y-auto">
              {addFiltered.length === 0 && (
                <div className="text-xs px-2 py-2" style={{ color: colors.inkFaint }}>{t.addMoreBatchesNoneMsg}</div>
              )}
              {addFiltered.slice(0, 30).map((i) => (
                <button
                  key={i.id}
                  type="button"
                  onClick={() => addBatchToDelivery(i.id)}
                  className="text-left px-2.5 py-1.5 rounded text-xs"
                  style={{ background: colors.surface, color: colors.ink, border: `1px solid ${colors.line}` }}
                >
                  <span className="font-semibold">{i.unitCode || i.id}</span> · {i.client} · {i.constructionSite || i.project}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {!form.recordedBy && (
        <div className="mt-4 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
          {t.recordedByRequiredMsg}
        </div>
      )}

      <div className="flex gap-2 mt-5">
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
          disabled={!anySelected || anyOvershoot || !form.date || !form.recordedBy}
          onClick={handleAddClick}
        >
          {isCombined ? t.addCombinedDeliveryBtn : t.addDeliveryBtn}
        </button>
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={onCancel}
        >
          {t.closeBtn}
        </button>
      </div>
    </div>
  );
}

const JOB_SHEET_TEMPLATES = ["Devan", "CFS", "Delivery", "Return", "Shifting", "Hoisting", "Day Work", "Dismantle", "Dis & Removal of Lifting Tools", "Job Cancel", "Pick-up", "Position", "Re-position", "Retain of Safety Ropes"];
// A Return brings cases back from site, so it is itemised and it adds to storage the
// same way an arrival does.
const JOB_SHEET_ITEMIZED = ["Devan", "CFS", "Delivery", "Return"];
const CFS_FROM_PRESETS = [
  { key: "金田物流", text: "金田物流有限公司\n油麻地海輝道38號新油麻地貨物公眾貨物\n裝卸區7號 3-4口水位高鴻躉船\n(車入閘后轉右直行約350米)\nTEL: 27850666 (亞昌/ 珍珍)" },
  { key: "悅昇物流", text: "悅昇物流有限公司\n香港新界葵涌昂船洲479號\nTEL: 24974884" },
];
const JOB_SHEET_BODY_SEEDS = {
  "Day Work": "- 按客戶要求將扶梯配件(共: _______箱) 由暫存位置攝出供客人開箱檢查貨件、\n  完成後用回原有包裝包妥，並攝回存倉位置擺放。",
  "Dismantle": "DISMANTLE WORK FOR ESCALATOR NO. ______\n\n1) 於下夜將所有吊重工具送交工地並移入所屬梯井。\n2) 將吊重工具掛上吊碼上。\n3) 將扶梯配件拆走。\n4) 將梯架吊離梯井。\n5) 用等鋰子切割機將扶梯切割至可移走呎吋。\n6) 機房內設置/搭建一套通架吊台供吊走摩打台。\n7) 出口外設置攝貨平台及圍封。\n8) 收走梯架及機房台。",
  "Dis & Removal of Lifting Tools": "FOR ESCALATOR NO. ______\n\n1/ 將有關工作台通架料於下夜送交工地及搬入______梯圍板內。\n2/ 於梯面搭建工作台將所有吊梯工具拆下並搬出站外上貨位。\n3/ 收離工地。",
  "Job Cancel": "- 因______，按______要求停止進行______工作。",
  "Pick-up": "- 客人安排運輸到快達倉取走全部______。(共: ______ 箱)",
  "Position": "POSITION WORK FOR ESCALATOR NO. ______\n\n1/ 設置吊重工具，以供吊嵌扶梯之用。\n2/ 吊臂車將扶梯及全部配件運到地盤。\n3/ 將扶梯攝入所屬梯井，並進行接駁、吊崁及安放到指定梯井位內。",
  "Re-position": "於______提供______(共______名)吊運人員將以上扶梯重新吊回梯井座好。",
  "Retain of Safety Ropes": "1) 於梯井設置羊眼圈及安全繩(共______組)供我司拆配件及吊走扶梯之用。\n2) 我司吊運工作完成後按客人要求保留梯井內(共______組)安全繩連羊眼圈供安裝工作之用。",
  "Shifting": "- 提供人手及工具將______攝到______暫存位置擺放。\n(TOTAL: ______ PKG / ______ KGS / ______ CBM)",
  "Hoisting": "HOISTING WORK FOR ______\n\n1/ ",
};

async function compressFileToDataUri(file, maxDim = 1400, quality = 0.72) {
  const readAsDataUrl = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(new Error("read-failed")); r.readAsDataURL(f); });
  if (file.type === "application/pdf") {
    if (file.size > 3.2 * 1024 * 1024) throw new Error("too-large");
    return await readAsDataUrl(file);
  }
  const dataUrl = await readAsDataUrl(file);
  const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error("bad-image")); i.src = dataUrl; });
  const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.width * scale));
  canvas.height = Math.max(1, Math.round(img.height * scale));
  canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/jpeg", quality);
}

function SignedDocControl({ docKey, colors, t }) {
  const [status, setStatus] = useState("loading");
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    let on = true;
    (async () => {
      try {
        const r = await storageGet(docKey);
        if (on) setStatus(r ? "saved" : "none");
      } catch (e) {
        if (on) setStatus("none");
      }
    })();
    return () => { on = false; };
  }, [docKey]);

  async function handleFile(e) {
    const f = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setBusy(true);
    try {
      const uri = await compressFileToDataUri(f);
      const ok = await storageSet(docKey, JSON.stringify({ uri, name: f.name, at: todayStr() }));
      if (!ok) throw new Error("save-failed");
      setStatus("saved");
    } catch (err) {
      alert(t.signedDocFailMsg);
    }
    setBusy(false);
  }

  async function view() {
    try {
      const r = await storageGet(docKey);
      if (!r) return;
      const { uri, name } = JSON.parse(r.value);
      const w = window.open("", "_blank");
      if (!w) return;
      if (uri.startsWith("data:application/pdf")) {
        w.document.write(`<title>${name || "Signed copy"}</title><embed src="${uri}" type="application/pdf" style="width:100%;height:100vh;">`);
      } else {
        w.document.write(`<title>${name || "Signed copy"}</title><body style="margin:0;background:#333;display:flex;justify-content:center;"><img src="${uri}" style="max-width:100%;height:auto;"></body>`);
      }
    } catch (e) { /* noop */ }
  }

  if (status === "loading") return <span className="text-xs" style={{ color: colors.inkFaint }}>…</span>;
  return (
    <span className="inline-flex items-center gap-1.5">
      {status === "saved" && (
        <button type="button" className="text-xs font-semibold underline" style={{ color: colors.green }} onClick={view}>
          ✓ {t.signedDocViewBtn}
        </button>
      )}
      <label className="text-xs font-semibold underline cursor-pointer" style={{ color: colors.amberText, opacity: busy ? 0.5 : 1 }}>
        {busy ? t.signedDocSavingMsg : status === "saved" ? t.signedDocReplaceBtn : t.signedDocUploadBtn}
        <input type="file" accept="image/*,application/pdf" style={{ display: "none" }} onChange={handleFile} disabled={busy} />
      </label>
    </span>
  );
}


// Real per-job Devan/Delivery handling tariffs, from the 2018 tariff sheets. This is a
// SEPARATE charge from the existing $/CBM/month storage billing - it's the cost of
// physically devanning/delivering the cargo, priced per job by R/Ton (revenue ton =
// max(weight in tons, volume in CBM)), varying by zone and elevator vs escalator.
// Only Chevalier and Schindler are populated - other clients keep using storage billing
// only until their tariffs are provided.
const HANDLING_TARIFFS = {
  Chevalier: {
    devanPerRTon: { urban: { elevator: 97, escalator: 97 }, lantau: { elevator: null, escalator: null } }, // Lantau devan: T.B.A. in tariff
    deliveryPerRTon: { urban: { elevator: 193, escalator: 193 }, lantau: { elevator: 232, escalator: 232 } },
    minPerTrip: { urban: 1920, lantau: 2300 },
    containerHaulage: { "20": 1550, "40": 1750 },
  },
  Schindler: {
    devanPerRTon: { urban: { elevator: 125, escalator: 125 }, lantau: { elevator: 125, escalator: 125 } },
    deliveryPerRTon: { urban: { elevator: 250, escalator: 250 }, lantau: { elevator: 300, escalator: 300 } },
    minPerTrip: null, // tariff doesn't give a plain min-per-trip separate from Special Minimum
    containerHaulage: { "20": 1550, "40": 1750 },
  },
};
function revenueTon(item) {
  const weightTons = (Number(item.weightKg) || 0) / 1000;
  const cbm = Number(item.volumeCbm) || (item.packages || []).reduce((s, p) => s + (Number(p.cbm) || 0), 0);
  return Math.max(weightTons, cbm);
}
// Computes the handling charge for one Devan or Delivery job on this item. Returns null
// if this client has no tariff on file yet. `amount: null` (with a rate table entry
// present but no number, e.g. Chevalier Lantau devan) means the tariff itself says
// "T.B.A." - needs a manual quote, not an app-computed price.
function computeHandlingCharge(item, jobType) {
  const tariff = HANDLING_TARIFFS[item.client];
  if (!tariff) return null;
  const zone = item.zone === "lantau" ? "lantau" : "urban";
  const cargoType = inferCargoType(item);
  const rTons = revenueTon(item);
  const rateTable = jobType === "devan" ? tariff.devanPerRTon : tariff.deliveryPerRTon;
  const perRTonRate = rateTable && rateTable[zone] ? rateTable[zone][cargoType] : undefined;
  if (perRTonRate == null) return { amount: null, rTons, zone, cargoType, needsQuote: true };
  let amount = rTons * perRTonRate;
  const minTrip = tariff.minPerTrip ? tariff.minPerTrip[zone] : null;
  if (minTrip && amount < minTrip) amount = minTrip;
  let oversizeMult = 1;
  if (item.isOversize && item.oversizeCbm) {
    oversizeMult = cargoType === "escalator"
      ? (oversizeLengthTierFor(item.client, Number(item.oversizeCbm)) || 1)
      : (oversizeTierFor(item.client, Number(item.oversizeCbm)) || 1);
  }
  amount *= oversizeMult;
  return { amount: Math.round(amount * 100) / 100, rTons, zone, cargoType, perRTonRate, minTrip, oversizeMult, needsQuote: false };
}
function computeContainerHaulageCharge(item) {
  const tariff = HANDLING_TARIFFS[item.client];
  if (!tariff) return null;
  const c20 = Number(item.containers20) || 0;
  const c40 = Number(item.containers40) || 0;
  if (!c20 && !c40) return null;
  const amount = c20 * tariff.containerHaulage["20"] + c40 * tariff.containerHaulage["40"];
  return { amount: Math.round(amount * 100) / 100, containers20: c20, containers40: c40 };
}
function inferCargoType(item) {
  if (item.cargoType) return item.cargoType;
  const text = [item.description, ...(item.packages || []).map((p) => p.description)].join(" ").toLowerCase();
  if (/escalator/.test(text)) return "escalator";
  if (/elevator|lift/.test(text)) return "elevator";
  return "elevator";
}

const OVERSIZE_RULES = {
  // Elevator packages, tiered by R/Ton (revenue ton = max(weight in tons, CBM)).
  // Anything above the highest tier isn't given an auto multiplier in the tariff itself -
  // it says "price to be advised" / quote separately, so oversizeTierFor returns null
  // (meaning: needs a manual quote) rather than guessing a rate.
  Schindler: [
    { min: 3.25, max: 5.50, mult: 1.5 },
    { min: 5.501, max: 7.50, mult: 2.0 },
  ],
  Chevalier: [
    { min: 4.0, max: 6.50, mult: 2.0 },
    { min: 6.501, max: 8.00, mult: 2.5 },
  ],
};
// Schindler's tariff separately tiers oversize ESCALATOR packages by length (metres),
// not R/Ton - Chevalier's tariff doesn't give escalator-specific oversize tiers at all.
const OVERSIZE_LENGTH_RULES = {
  Schindler: [
    { min: 7.901, max: 13.5, mult: 2.0 },
  ],
};
function oversizeLengthTierFor(client, meters) {
  const rules = OVERSIZE_LENGTH_RULES[client];
  if (!rules || !(meters > 0)) return null;
  for (const r of rules) {
    if (meters >= r.min && meters <= r.max) return r.mult;
  }
  return null; // above the highest tier (or below it) - not auto-priced, quote separately
}
function oversizeTierFor(client, cbm) {
  const rules = OVERSIZE_RULES[client];
  if (!rules || !(cbm > 0)) return null;
  for (const r of rules) {
    if (cbm >= r.min && cbm <= r.max) return r.mult;
  }
  return null;
}
function computeOversizeText(item) {
  const rules = OVERSIZE_RULES[item.client];
  if (!rules || !(item.packages || []).length) return "";
  const byMult = new Map();
  for (const p of item.packages) {
    const cbm = Number(p.cbm);
    const mult = oversizeTierFor(item.client, cbm);
    if (!mult) continue;
    if (!byMult.has(mult)) byMult.set(mult, []);
    byMult.get(mult).push({ code: p.code, cbm });
  }
  if (byMult.size === 0) return "";
  const lines = [];
  const sortedMults = [...byMult.keys()].sort((a, b) => a - b);
  for (const mult of sortedMults) {
    const entries = byMult.get(mult);
    const byCbm = new Map();
    for (const e of entries) {
      const key = String(e.cbm);
      if (!byCbm.has(key)) byCbm.set(key, []);
      byCbm.get(key).push(e.code);
    }
    for (const [cbmVal, codes] of byCbm) {
      lines.push(`#${codes.join(",")} @${cbmVal}CBM`);
    }
    const subtotal = Math.round(entries.reduce((s, e) => s + e.cbm, 0) * 100) / 100;
    lines.push(`(合共 Total: ${subtotal} CBM)  x${mult} 倍`);
    lines.push("");
  }
  return lines.join("\n").trim();
}

function deliveryTotalsFor(it, del) {
  const pkgs = del.codes ? del.codes.length : Number(del.packageCount) || 0;
  let kgs = "", cbm = "", estimated = false;
  if (del.codes && (it.packages || []).length > 0) {
    const delivered = it.packages.filter((p) => del.codes.includes(p.code));
    const haveAllWeights = delivered.length > 0 && delivered.every((p) => p.weightKg !== "" && p.weightKg != null);
    const haveAllCbm = delivered.length > 0 && delivered.every((p) => p.cbm !== "" && p.cbm != null);
    kgs = haveAllWeights ? String(Math.round(delivered.reduce((s, p) => s + Number(p.weightKg), 0) * 10) / 10) : "";
    cbm = haveAllCbm ? String(Math.round(delivered.reduce((s, p) => s + Number(p.cbm), 0) * 1000) / 1000) : "";
    if (!haveAllWeights || !haveAllCbm) estimated = true;
  } else {
    estimated = true;
  }
  if (estimated) {
    const totalU = totalUnits(it) || 1;
    const share = pkgs / totalU;
    if (!kgs && it.weightKg) kgs = `~${Math.round(Number(it.weightKg) * share * 10) / 10}`;
    if (!cbm && it.volumeCbm) cbm = `~${Math.round(Number(it.volumeCbm) * share * 1000) / 1000}`;
  }
  return { pkgs, kgs, cbm, estimated };
}
function refTextFor(it, del) {
  const jn = it.jobNumber || "";
  let dates = [];
  if (usesArrivalBatches(it) && del && del.codes) {
    const set = new Set(del.codes);
    dates = [...new Set(activeArrivals(it).filter((a) => (a.codes || []).some((c) => set.has(c))).map((a) => a.date).filter(Boolean))];
  }
  if (dates.length === 0) dates = [effectiveDepotArrivalDate(it)].filter(Boolean);
  if (dates.length === 0 || !jn) return "";
  return dates.map((d) => `Refer to job no. ${jn} on ${fmt(d)}`).join("\n");
}

function CombinedDeliveryPrint({ sheet, onClose, directory, colors, t, lang }) {
  const { groups, jobNumber, date, deliveredTo } = sheet;
  const firstItem = groups[0].item;
  const depotEn = firstItem.depot || "";
  const depotZh = DEPOT_LABELS_ZH[firstItem.depot] || "";
  const dirMatch = (directory || []).find((s) => s.siteEn && (s.siteEn === firstItem.constructionSite || s.siteEn === firstItem.project));
  const siteZh = (dirMatch && dirMatch.siteZh) || "";
  const siteContact = dirMatch ? [dirMatch.contactName, dirMatch.contactPhone].filter(Boolean).join(" ") : "";

  const site = firstItem.constructionSite || firstItem.project || deliveredTo || "";
  const toTop = [site ? (/^site\s+at/i.test(site) ? site : `SITE AT ${site}`) : "", siteZh].filter(Boolean).join("\n");
  const toBottom = siteContact;

  const rows = groups.map((g) => ({ ...g, totals: deliveryTotalsFor(g.item, g.delivery), ref: refTextFor(g.item, g.delivery) }));
  const anyEstimated = rows.some((r) => r.totals.estimated);
  const totalPkgs = rows.reduce((s, r) => s + r.totals.pkgs, 0);
  const numOrNull = (v) => { const n = Number(String(v || "").replace("~", "")); return isNaN(n) ? null : n; };
  const kgVals = rows.map((r) => numOrNull(r.totals.kgs));
  const cbmVals = rows.map((r) => numOrNull(r.totals.cbm));
  const totalKgs = kgVals.every((v) => v != null) ? Math.round(kgVals.reduce((s, v) => s + v, 0) * 10) / 10 : "";
  const totalCbm = cbmVals.every((v) => v != null) ? Math.round(cbmVals.reduce((s, v) => s + v, 0) * 1000) / 1000 : "";

  const lbl = { border: "1px solid #111", fontWeight: "bold", padding: "4px 6px", verticalAlign: "top", width: 78, fontSize: 13, wordBreak: "break-word", overflowWrap: "break-word" };
  const cel = { border: "1px solid #111", padding: 6, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };
  const cel0 = { border: "1px solid #111", padding: 0, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };

  return (
    <div id="job-sheet-print-root" className="fixed inset-0 z-50 flex flex-col" style={{ background: "rgba(0,0,0,0.5)" }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #job-sheet-print-root, #job-sheet-print-root * { visibility: visible; }
          #job-sheet-print-root { position: absolute; inset: auto; left: 0; top: 0; width: 100%; background: #fff !important; }
          #job-sheet-print-toolbar { display: none !important; }
          #job-sheet-print-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          #job-sheet-print-area { margin: 0 !important; max-width: 100% !important; page-break-inside: avoid; break-inside: avoid; }
        }
        @page { size: A4 portrait; margin: 10mm; }
      `}</style>
      <div id="job-sheet-print-toolbar" className="flex flex-wrap items-center gap-2 p-3" style={{ background: colors.navy }}>
        <span className="text-sm font-semibold" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>{t.combinedPrintLabel}</span>
        <div className="ml-auto flex gap-2">
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => window.print()}>
            {t.printBtn}
          </button>
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.onDark}`, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>
            {t.closePreviewBtn}
          </button>
        </div>
      </div>
      <div id="job-sheet-print-scroll" className="flex-1 overflow-y-auto p-6" style={{ background: colors.bg }}>
        <div id="job-sheet-print-area" className="mx-auto" style={{ background: "#fff", color: "#111", maxWidth: 700, padding: 24, fontFamily: "Arial, sans-serif", fontSize: 15 }}>

          <div className="flex items-start justify-between" style={{ marginBottom: 8 }}>
            <div className="flex items-center gap-3">
              <img src={FARSPEED_LOGO_DATA_URI} alt="Farspeed" style={{ height: 78, width: "auto" }} />
              <div>
                <div style={{ fontSize: 24, fontWeight: 700, color: "#111", lineHeight: 1.1 }}>FARSPEED Contractors Limited</div>
                <div style={{ fontSize: 11, color: "#111" }}>P. O. Box No. 1985, Yuen Long Post Office, Yuen Long, N.T., Hong Kong</div>
                <div style={{ fontSize: 11, color: "#111" }}>Tel: +852 5337-9500&nbsp;&nbsp;Fax: +852 2402-4450&nbsp;&nbsp;http://www.farspeed.hk</div>
              </div>
            </div>
            <div className="flex items-start" style={{ fontSize: 16, fontWeight: 700, lineHeight: 1.5 }}>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3 }}>有限公司</div>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3, marginLeft: 4 }}>快達承判</div>
            </div>
          </div>
          <div style={{ borderTop: "2.5px solid #111", marginBottom: 10 }} />

          <div className="text-center font-bold mb-3" style={{ fontSize: 22, letterSpacing: 6 }}>
            {t.jsTitleZh}&nbsp;&nbsp;{t.jsTitle}
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18 }}>
            <tbody>
              <tr>
                <td style={lbl}>{t.jsFromZh}<br />{t.jsFrom}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: "1px solid #111" }}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20 }}>{[depotZh, depotEn].filter(Boolean).join("\n")}</div>
                </td>
                <td style={lbl} rowSpan={toBottom ? 1 : 1}>{t.jsToZh}<br />{t.jsTo}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: toBottom ? "none" : "1px solid #111" }} rowSpan={toBottom ? 1 : 1}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{toTop}</div>
                </td>
              </tr>
              {toBottom && (
                <tr>
                  <td style={lbl}></td>
                  <td style={cel0}></td>
                  <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderBottom: "1px solid #111", borderTop: "none" }}>
                    <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{toBottom}</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>

          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18, marginTop: -1 }}>
            <tbody>
              <tr>
                <td style={lbl}>{t.jsAccountZh}<br />{t.jsAccount}</td>
                <td style={cel}>{firstItem.client}</td>
                <td style={lbl}>{t.jsJobNoZh}<br />{t.jsJobNo}</td>
                <td style={{ ...cel, fontWeight: "bold" }}>{jobNumber || "—"}</td>
                <td style={lbl}>{t.jsDateZh}<br />{t.jsDate}</td>
                <td style={cel}>{fmt(date)}</td>
              </tr>
              <tr>
                <td style={lbl}>{t.jsOrderedByZh}<br />{t.jsOrderedBy}</td>
                <td style={cel}>{firstItem.orderedBy || "—"}</td>
                <td style={lbl}>{t.jsPoNoZh}<br />{t.jsPoNo}</td>
                <td style={cel}>{firstItem.poNumber || "—"}</td>
                <td style={lbl}>{t.jsJobRefZh}<br />{t.jsJobRef}</td>
                <td style={cel}>{firstItem.jobRef || "—"}</td>
              </tr>
              <tr>
                <td style={lbl}>提單資料<br />SS/D.O. NO.</td>
                <td style={cel} colSpan={5}>—</td>
              </tr>
            </tbody>
          </table>

          <div style={{ border: "1px solid #111", borderTop: "none", padding: "4px 6px", display: "flex", alignItems: "center", fontSize: 16 }}>
            <span style={{ flex: 1, textAlign: "center", fontWeight: "bold" }}>
              {t.jsDescriptionZh}<br /><span style={{ letterSpacing: 2 }}>{t.jsDescription}</span>
            </span>
            <span style={{ fontSize: 12 }}>{t.jsIssuedByZh}: {groups[0].delivery.recordedBy || "—"}</span>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", border: "1px solid #111", borderTop: "none", fontSize: 16 }}>
            <tbody>
              <tr>
                <td style={{ padding: 8, verticalAlign: "top", height: 400 }}>
                  {rows.map((r, idx) => (
                    <div key={r.item.id} style={{ marginBottom: idx < rows.length - 1 ? 14 : 0 }}>
                      {r.ref && <div style={{ marginBottom: 4, textDecoration: "underline", whiteSpace: "pre-line" }}>{r.ref}</div>}
                      {r.item.shkNumber && <div style={{ fontWeight: "bold" }}>{r.item.shkNumber}</div>}
                      {r.item.unitCode && <div style={{ fontWeight: "bold" }}>{r.item.unitCode}</div>}
                      {r.item.description && <div>{r.item.description}</div>}
                      {r.delivery.codes && r.delivery.codes.length > 0 && (
                        <div style={{ fontSize: 12, color: "#111" }}>
                          C/S NO. {r.delivery.codes.join(", ")} &nbsp;&nbsp; {r.totals.pkgs} {t.jsPkgs}
                          {r.totals.kgs ? ` \u00b7 ${r.totals.kgs} ${t.jsKgs}` : ""}{r.totals.cbm ? ` \u00b7 ${r.totals.cbm} ${t.jsCbm}` : ""}
                        </div>
                      )}
                    </div>
                  ))}
                  <div style={{ borderTop: "1px solid #111", marginTop: 12, paddingTop: 6 }}>
                    共:&nbsp;&nbsp;&nbsp;{totalPkgs} {t.jsPkgs} &nbsp;&nbsp;&nbsp; {totalKgs !== "" ? totalKgs : "—"} {t.jsKgs} &nbsp;&nbsp;&nbsp; {totalCbm !== "" ? totalCbm : "—"} {t.jsCbm}
                  </div>
                  {anyEstimated && <div style={{ fontSize: 10, color: "#900", marginTop: 4 }}>{t.jsEstimatedNote}</div>}
                </td>
              </tr>
              <tr>
                <td style={{ padding: "0 8px", fontSize: 16 }}>
                  客戶簽署確認 / Customer&nbsp;signature:&nbsp;______________________________&nbsp;(工作妥當及完成)
                  <br /><br /><br /><br />
                </td>
              </tr>
            </tbody>
          </table>

          <div style={{ borderTop: "1px solid #ccc", paddingTop: 5, fontSize: 8, color: "#111", textAlign: "center", marginTop: 8 }}>
            Office and Depot: 21D Wang Toi Shan, Hung Mo Tam, Kam Tin. NT., HK
            <div style={{ marginTop: 2 }}>N.B. Farspeed Contractors Ltd. is a private company. All transaction(s) taken into account are subject to the STANDARD BUSINESS CONDITIONS of the company, details as behind.</div>
            <div style={{ marginTop: 2 }}>(a member of FARSPEED Group)</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function JobSheetPrint({ sheet, onClose, directory, colors, t, lang }) {
  if (sheet.combined) {
    return <CombinedDeliveryPrint sheet={sheet} onClose={onClose} directory={directory} colors={colors} t={t} lang={lang} />;
  }
  const { type, item, delivery } = sheet;
  const isDelivery = type === "Delivery";
  const depotEn = item.depot || "";
  const depotZh = DEPOT_LABELS_ZH[item.depot] || "";
  const dirMatch = (directory || []).find((s) => s.siteEn && (s.siteEn === item.constructionSite || s.siteEn === item.project));
  const siteZh = (dirMatch && dirMatch.siteZh) || "";
  const siteContact = dirMatch ? [dirMatch.contactName, dirMatch.contactPhone].filter(Boolean).join(" ") : "";

  const initialTemplate = isDelivery ? "Delivery" : (JOB_SHEET_TEMPLATES.includes(type) ? type : "Devan");
  const [template, setTemplate] = useState(initialTemplate);
  const [cfsPresetKey, setCfsPresetKey] = useState(CFS_FROM_PRESETS[0].key);

  // The Delivery sheet names the depot in both languages - 快達一號倉 / Farspeed Depot 1 -
  // and Devan and CFS only did so in Chinese. Same treatment across all three now.
  function fromTopText(tpl) {
    if (tpl === "Devan") {
      const zh = `客人安排運輸送到${depotZh || depotEn} (共1櫃)`;
      const en = depotEn ? `Customer-arranged delivery to ${depotEn} (1 container)` : "";
      return [zh, en].filter(Boolean).join("\n");
    }
    if (tpl === "CFS") return (CFS_FROM_PRESETS.find((p) => p.key === cfsPresetKey) || CFS_FROM_PRESETS[0]).text;
    if (["Delivery", "Pick-up", "Day Work"].includes(tpl)) return [depotZh, depotEn].filter(Boolean).join("\n");
    return "";
  }
  function fromBottomText(tpl) {
    if (tpl === "Devan") return "*由快達拆櫃 / Devanned by Farspeed";
    if (tpl === "CFS") return "*由客戶自行CFS / Customer's own CFS";
    return "";
  }
  function toTopText() {
    const site = item.constructionSite || item.project || "";
    const out = [];
    if (site) out.push(/^site\s+at/i.test(site) ? site : `SITE AT ${site}`);
    if (siteZh) out.push(siteZh);
    return out.join("\n");
  }
  function toBottomText(tpl) {
    if (tpl === "Devan" || tpl === "CFS") {
      return depotEn ? `暫存${depotZh || depotEn} / Stored at ${depotEn}` : `暫存${depotZh || depotEn}`;
    }
    if (tpl === "Delivery") return siteContact || "";
    return "";
  }
  function ssText(tpl) {
    if (tpl === "Devan" || tpl === "CFS") return item.ssDoNo || "";
    return "";
  }
  function refText(tpl) {
    if (tpl !== "Delivery") return "";
    const jn = item.jobNumber || "";
    let dates = [];
    if (usesArrivalBatches(item) && delivery && delivery.codes) {
      const set = new Set(delivery.codes);
      dates = [...new Set(activeArrivals(item).filter((a) => (a.codes || []).some((c) => set.has(c))).map((a) => a.date).filter(Boolean))];
    }
    if (dates.length === 0) dates = [effectiveDepotArrivalDate(item)].filter(Boolean);
    if (dates.length === 0 || !jn) return "";
    return dates.map((d) => `Refer to job no. ${jn} on ${fmt(d)}`).join("\n");
  }
  function bodyText(tpl) {
    if (JOB_SHEET_ITEMIZED.includes(tpl)) return "";
    const site = (item.constructionSite || item.project || "").trim();
    const seed = JOB_SHEET_BODY_SEEDS[tpl] || "";
    return [site, seed].filter(Boolean).join("\n\n");
  }

  const itemized = JOB_SHEET_ITEMIZED.includes(template);
  const dateText = isDelivery ? delivery.date : item.depotArrivalDate || effectiveDepotArrivalDate(item);
  const jobNo = isDelivery ? delivery.jobNumber : item.jobNumber;
  const issuedBy = (isDelivery ? delivery.recordedBy : item.recordedBy) || "";
  // A delivery sheet carries its own OVERSIZE CASES box - the cases going out on this
  // job, named case by case ("L13#13/23@4.49CBM") - so it prints what was recorded
  // against the delivery rather than re-deriving the whole lot's oversize cases.
  const deliveryOversizeCases = isDelivery && delivery && delivery.oversize
    ? cleanOversizeCases(delivery.oversize.cases)
    : [];
  const showOversize = deliveryOversizeCases.length > 0
    ? true
    : (["Devan", "CFS"].includes(template) && !!OVERSIZE_RULES[item.client]);
  const oversizeText = deliveryOversizeCases.length > 0
    ? deliveryOversizeCases.map((c) => `${c.code ? `#${c.code} ` : ""}@${c.cbm}CBM`).join("\n")
    : (showOversize ? computeOversizeText(item) : "");

  const pkgs = isDelivery ? (delivery.codes ? delivery.codes.length : Number(delivery.packageCount) || 0) : totalUnits(item);
  let kgs = item.weightKg || "";
  let cbm = item.volumeCbm || "";
  let estimated = false;
  if (isDelivery) {
    // A delivery recorded from a job sheet carries that sheet's own totals. They are the
    // figures the job was done on, so the printed sheet states them rather than a fresh
    // count of per-case packing-list weights that would disagree with the paperwork.
    const decKg = delivery.declared ? declaredNum(delivery.declared.kg) : null;
    const decCbm = delivery.declared ? declaredNum(delivery.declared.cbm) : null;
    if (delivery.codes && (item.packages || []).length > 0) {
      const delivered = item.packages.filter((p) => delivery.codes.includes(p.code));
      const haveAllWeights = delivered.length > 0 && delivered.every((p) => p.weightKg !== "" && p.weightKg != null);
      const haveAllCbm = delivered.length > 0 && delivered.every((p) => p.cbm !== "" && p.cbm != null);
      kgs = haveAllWeights ? String(Math.round(delivered.reduce((s, p) => s + Number(p.weightKg), 0) * 10) / 10) : "";
      cbm = haveAllCbm ? String(Math.round(delivered.reduce((s, p) => s + Number(p.cbm), 0) * 1000) / 1000) : "";
      if (!haveAllWeights || !haveAllCbm) estimated = true;
    } else {
      estimated = true;
    }
    if (estimated) {
      const totalU = totalUnits(item) || 1;
      const share = pkgs / totalU;
      if (!kgs && item.weightKg) kgs = `~${Math.round(Number(item.weightKg) * share * 10) / 10}`;
      if (!cbm && item.volumeCbm) cbm = `~${Math.round(Number(item.volumeCbm) * share * 1000) / 1000}`;
    }
    // Stated on the job sheet beats anything derived, and is not an estimate.
    if (decKg != null) kgs = String(decKg);
    if (decCbm != null) cbm = String(decCbm);
    if (decKg != null && decCbm != null) estimated = false;
  }
  const csLine = isDelivery
    ? (delivery.codes ? delivery.codes.join(", ") : "")
    : (item.packages || []).map((p) => p.code).join(", ");

  const lbl = { border: "1px solid #111", fontWeight: "bold", padding: "4px 6px", verticalAlign: "top", width: 78, fontSize: 13, wordBreak: "break-word", overflowWrap: "break-word" };
  const cel = { border: "1px solid #111", padding: 6, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };
  const cel0 = { border: "1px solid #111", padding: 0, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };
  const fTop = fromTopText(template), fBottom = fromBottomText(template);
  const tTop = toTopText(), tBottom = toBottomText(template);
  const fromHasBottom = !!fBottom, toHasBottom = !!tBottom;
  const bText = bodyText(template), rText = refText(template), sText = ssText(template);
  // The description box was a flat 400px, which suited Delivery's compact two-line
  // FROM/TO block but pushed the signature line and footer onto a second page on Devan,
  // and further still on CFS where the forwarder's address runs seven lines. Give the box
  // whatever height is left instead, so every template prints on a single page.
  const lineCount = (s) => (s ? String(s).split("\n").length : 0);
  const fromToLines = Math.max(lineCount(fTop) + lineCount(fBottom), lineCount(tTop) + lineCount(tBottom));
  // Everything above and below the description box is a known height; the box gets what's
  // left of the page rather than a flat 400px. A Delivery's compact FROM/TO block leaves
  // exactly the 400 it always had, so that sheet is unchanged. A CFS, whose forwarder
  // address runs five or six lines, and a Devan, which also prints an oversize block,
  // give some of it back instead of pushing the signature onto a second page.
  const PAGE_CONTENT_PX = 1020;  // A4 at 96dpi less 10mm margins
  const FIXED_CHROME_PX = 520;   // letterhead, title, header table, description bar, signature, footer
  const oversizePx = showOversize && oversizeText ? 40 + lineCount(oversizeText) * 22 : 0;
  const descHeight = Math.max(200, Math.min(400, PAGE_CONTENT_PX - FIXED_CHROME_PX - (40 + fromToLines * 30) - oversizePx));

  return (
    <div id="job-sheet-print-root" className="fixed inset-0 z-50 flex flex-col" style={{ background: "rgba(0,0,0,0.5)" }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #job-sheet-print-root, #job-sheet-print-root * { visibility: visible; }
          #job-sheet-print-root { position: absolute; inset: auto; left: 0; top: 0; width: 100%; background: #fff !important; }
          #job-sheet-print-toolbar { display: none !important; }
          #job-sheet-print-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          #job-sheet-print-area { margin: 0 !important; max-width: 100% !important; page-break-inside: avoid; break-inside: avoid; }
        }
        @page { size: A4 portrait; margin: 10mm; }
      `}</style>
      <div id="job-sheet-print-toolbar" className="flex flex-wrap items-center gap-2 p-3" style={{ background: colors.navy }}>
        <span className="text-sm font-semibold" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>{t.jsTemplateLabel}:</span>
        <select
          value={template}
          onChange={(e) => setTemplate(e.target.value)}
          className="px-2 py-1.5 rounded text-sm"
          style={{ background: colors.surface, color: colors.ink }}
        >
          {JOB_SHEET_TEMPLATES.map((tp) => <option key={tp} value={tp}>{tp}</option>)}
        </select>
        {template === "CFS" && (
          <select
            value={cfsPresetKey}
            onChange={(e) => setCfsPresetKey(e.target.value)}
            className="px-2 py-1.5 rounded text-sm"
            style={{ background: colors.surface, color: colors.ink }}
          >
            {CFS_FROM_PRESETS.map((p) => <option key={p.key} value={p.key}>{p.key}</option>)}
          </select>
        )}
        <div className="ml-auto flex gap-2">
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => window.print()}>
            {t.printBtn}
          </button>
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.onDark}`, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>
            {t.closePreviewBtn}
          </button>
        </div>
      </div>
      <div id="job-sheet-print-scroll" className="flex-1 overflow-y-auto p-6" style={{ background: colors.bg }}>
        <div id="job-sheet-print-area" className="mx-auto" style={{ background: "#fff", color: "#111", maxWidth: 700, padding: 24, fontFamily: "Arial, sans-serif", fontSize: 15 }}>

          {/* Letterhead */}
          <div className="flex items-start justify-between" style={{ marginBottom: 8 }}>
            <div className="flex items-center gap-3">
              <img src={FARSPEED_LOGO_DATA_URI} alt="Farspeed" style={{ height: 78, width: "auto" }} />
              <div>
                <div style={{ fontSize: 24, fontWeight: 700, color: "#111", lineHeight: 1.1 }}>FARSPEED Contractors Limited</div>
                <div style={{ fontSize: 11, color: "#111" }}>P. O. Box No. 1985, Yuen Long Post Office, Yuen Long, N.T., Hong Kong</div>
                <div style={{ fontSize: 11, color: "#111" }}>Tel: +852 5337-9500&nbsp;&nbsp;Fax: +852 2402-4450&nbsp;&nbsp;http://www.farspeed.hk</div>
              </div>
            </div>
            <div className="flex items-start" style={{ fontSize: 16, fontWeight: 700, lineHeight: 1.5 }}>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3 }}>有限公司</div>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3, marginLeft: 4 }}>快達承判</div>
            </div>
          </div>
          <div style={{ borderTop: "2.5px solid #111", marginBottom: 10 }} />

          <div className="text-center font-bold mb-3" style={{ fontSize: 22, letterSpacing: 6 }}>
            {t.jsTitleZh}&nbsp;&nbsp;{t.jsTitle}
          </div>

          {/* FROM / TO — one real 2-row table so the two sides line up on the same horizontal, no divider between rows */}
          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18 }}>
            <tbody>
              <tr>
                <td style={lbl} rowSpan={2}>{t.jsFromZh}<br />{t.jsFrom}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: fromHasBottom ? "none" : "1px solid #111" }} rowSpan={fromHasBottom ? 1 : 2}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20 }}>{fTop}</div>
                </td>
                <td style={lbl} rowSpan={2}>{t.jsToZh}<br />{t.jsTo}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: toHasBottom ? "none" : "1px solid #111" }} rowSpan={toHasBottom ? 1 : 2}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{tTop}</div>
                </td>
              </tr>
              <tr>
                {fromHasBottom && (
                  <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderBottom: "1px solid #111", borderTop: "none" }}>
                    <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{fBottom}</div>
                  </td>
                )}
                {toHasBottom && (
                  <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderBottom: "1px solid #111", borderTop: "none" }}>
                    <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{tBottom}</div>
                  </td>
                )}
              </tr>
            </tbody>
          </table>

          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18, marginTop: -1 }}>
            <tbody>
              <tr>
                <td style={lbl}>{t.jsAccountZh}<br />{t.jsAccount}</td>
                <td style={cel}>{item.client}</td>
                <td style={lbl}>{t.jsJobNoZh}<br />{t.jsJobNo}</td>
                <td style={{ ...cel, fontWeight: "bold" }}>{jobNo || "—"}</td>
                <td style={lbl}>{t.jsDateZh}<br />{t.jsDate}</td>
                <td style={cel}>{fmt(dateText)}</td>
              </tr>
              <tr>
                <td style={lbl}>{t.jsOrderedByZh}<br />{t.jsOrderedBy}</td>
                <td style={cel}>{item.orderedBy || "—"}</td>
                <td style={lbl}>{t.jsPoNoZh}<br />{t.jsPoNo}</td>
                <td style={cel}>{item.poNumber || "—"}</td>
                <td style={lbl}>{t.jsJobRefZh}<br />{t.jsJobRef}</td>
                <td style={cel}>{item.jobRef || "—"}</td>
              </tr>
              <tr>
                <td style={lbl}>提單資料<br />SS/D.O. NO.</td>
                <td style={cel} colSpan={5}>{sText || "—"}</td>
              </tr>
            </tbody>
          </table>

          <div style={{ border: "1px solid #111", borderTop: "none", padding: "4px 6px", display: "flex", alignItems: "center", fontSize: 16 }}>
            <span style={{ flex: 1, textAlign: "center", fontWeight: "bold" }}>
              {t.jsDescriptionZh}<br /><span style={{ letterSpacing: 2 }}>{t.jsDescription}</span>
            </span>
            <span style={{ fontSize: 12 }}>{t.jsIssuedByZh}: {issuedBy || "—"}</span>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", border: "1px solid #111", borderTop: "none", fontSize: 16 }}>
            <tbody>
              <tr>
                <td style={{ padding: 8, verticalAlign: "top", height: descHeight }}>
                  {itemized ? (
                    <>
                      {template === "Delivery" && rText && <div style={{ marginBottom: 6, textDecoration: "underline", whiteSpace: "pre-line" }}>{rText}</div>}
                      {["Devan", "CFS", "Delivery"].includes(template) && item.shkNumber && <div style={{ fontWeight: "bold" }}>{item.shkNumber}</div>}
                      {item.unitCode && <div style={{ fontWeight: "bold" }}>{item.unitCode}</div>}
                      {item.description && <div>{item.description}</div>}
                      {csLine && <div style={{ fontSize: 12, color: "#111" }}>C/S NO. {csLine}</div>}
                      <div style={{ borderTop: "1px solid #111", marginTop: 12, paddingTop: 6 }}>
                        共:&nbsp;&nbsp;&nbsp;{pkgs} {t.jsPkgs} &nbsp;&nbsp;&nbsp; {kgs || "—"} {t.jsKgs} &nbsp;&nbsp;&nbsp; {cbm || "—"} {t.jsCbm}
                      </div>
                      {estimated && <div style={{ fontSize: 10, color: "#900", marginTop: 4 }}>{t.jsEstimatedNote}</div>}
                      {showOversize && oversizeText && (
                        <div style={{ borderTop: "1px solid #111", marginTop: 12, paddingTop: 8 }}>
                          <div style={{ fontWeight: "bold", textDecoration: "underline" }}>{t.jsOversizeLabel}</div>
                          <div style={{ marginTop: 4, fontSize: 13, whiteSpace: "pre-line" }}>{oversizeText}</div>
                          <div style={{ fontSize: 10, color: "#111", marginTop: 4 }}>{t.jsOversizeNote}</div>
                        </div>
                      )}
                    </>
                  ) : (
                    <div style={{ whiteSpace: "pre-line" }}>{bText}</div>
                  )}
                </td>
              </tr>
              <tr>
                <td style={{ padding: "0 8px", fontSize: 16 }}>
                  {template === "Pick-up" ? (
                    <>
                      簽署確認收妥 / Signature:&nbsp;______________________________&nbsp;(車輛號碼 Vehicle&nbsp;No.:________ )
                      <br /><br /><br />
                    </>
                  ) : (
                    <>
                      客戶簽署確認 / Customer&nbsp;signature:&nbsp;______________________________&nbsp;(工作妥當及完成)
                      <br /><br /><br /><br />
                    </>
                  )}
                </td>
              </tr>
            </tbody>
          </table>

          <div style={{ borderTop: "1px solid #ccc", paddingTop: 5, fontSize: 8, color: "#111", textAlign: "center", marginTop: 8 }}>
            Office and Depot: 21D Wang Toi Shan, Hung Mo Tam, Kam Tin. NT., HK
            <div style={{ marginTop: 2 }}>N.B. Farspeed Contractors Ltd. is a private company. All transaction(s) taken into account are subject to the STANDARD BUSINESS CONDITIONS of the company, details as behind.</div>
            <div style={{ marginTop: 2 }}>(a member of FARSPEED Group)</div>
          </div>
        </div>
      </div>
    </div>
  );
}

// --- Debit notes -----------------------------------------------------------------
// Farspeed bills a client per site per month, and the debit note that goes out is typed
// up outside this app. Recording it here lets the amount charged be set against the
// amount the depot's own records produce, so a wrong figure shows up before the client
// finds it. Hoisting and Shifting are deliberately absent - they are a later job.
const DEBIT_NOTE_CATEGORIES = ["Storage", "CFS", "Devan", "Delivery"];
const NUM_WORDS_ONES = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN",
  "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"];
const NUM_WORDS_TENS = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"];
function numberToWordsUnder1000(n) {
  const out = [];
  if (n >= 100) { out.push(`${NUM_WORDS_ONES[Math.floor(n / 100)]} HUNDRED`); n %= 100; }
  if (n >= 20) {
    const tens = NUM_WORDS_TENS[Math.floor(n / 10)];
    const ones = NUM_WORDS_ONES[n % 10];
    out.push(ones ? `${tens}-${ones}` : tens);
  } else if (n > 0) out.push(NUM_WORDS_ONES[n]);
  return out.join(" ");
}
// "SAY TOTAL HONG KONG DOLLARS FOUR THOUSAND FIVE HUNDRED NINETY-FIVE AND ELEVEN CENTS ONLY."
function amountInWordsHKD(amount) {
  const value = Math.round((Number(amount) || 0) * 100) / 100;
  const dollars = Math.floor(value);
  const cents = Math.round((value - dollars) * 100);
  const parts = [];
  let rest = dollars;
  const scales = [[1000000000, "BILLION"], [1000000, "MILLION"], [1000, "THOUSAND"]];
  for (const [size, name] of scales) {
    if (rest >= size) {
      parts.push(`${numberToWordsUnder1000(Math.floor(rest / size))} ${name}`);
      rest %= size;
    }
  }
  if (rest > 0) parts.push(numberToWordsUnder1000(rest));
  if (!parts.length) parts.push("ZERO");
  const words = `SAY TOTAL HONG KONG DOLLARS ${parts.join(" ")}`;
  return cents > 0
    ? `${words} AND ${numberToWordsUnder1000(cents)} CENTS ONLY.`
    : `${words} ONLY.`;
}
// What the depot's own records say this client should be charged for one site, one
// category and one month. Storage comes off the per-case storage clock; CFS, Devan and
// Delivery come off the handling tariff. Returns the lines behind the figure too, so a
// mismatch can be looked into rather than just reported.
function computeCategoryTotal(items, { client, project, category, year, month }) {
  const lines = [];
  const onSite = (item) => !project
    || sitesLooselyMatch(project, "", item.project, item.constructionSite)
    || sitesLooselyMatch(project, "", item.constructionSite, item.project);
  const inPeriod = (dateStr) => {
    if (!dateStr) return false;
    const d = toDateOnly(dateStr);
    return d.getFullYear() === year && d.getMonth() === month;
  };
  for (const item of items || []) {
    if (item.cancelled) continue;
    if (client && item.client !== client) continue;
    if (!onSite(item)) continue;
    if (category === "Storage") {
      for (const row of computeItemBillingRows(item)) {
        for (const line of row.breakdown) {
          if (line.year !== year || line.month !== month) continue;
          lines.push({
            itemId: item.id, jobNumber: item.jobNumber || "", unitCode: item.unitCode || "",
            detail: line.detail || line.label || "", amount: Math.round(line.amount * 100) / 100,
            estimated: !!row.estimated,
          });
        }
      }
      continue;
    }
    if (category === "Devan" || category === "CFS") {
      // The arrival charge is the same tariff either way; which of the two it is depends on
      // how the cases were checked in.
      if ((item.arrivingType || "") !== category) continue;
      if (!inPeriod(effectiveDepotArrivalDate(item))) continue;
      const c = computeHandlingCharge(item, "devan");
      if (c && c.amount) lines.push({
        itemId: item.id, jobNumber: item.jobNumber || "", unitCode: item.unitCode || "",
        detail: c.detail || category, amount: Math.round(c.amount * 100) / 100, needsQuote: !!c.needsQuote,
      });
      continue;
    }
    if (category === "Delivery") {
      if (!inPeriod(lastDeliveryDate(item))) continue;
      const c = computeHandlingCharge(item, "delivery");
      if (c && c.amount) lines.push({
        itemId: item.id, jobNumber: item.jobNumber || "", unitCode: item.unitCode || "",
        detail: c.detail || "Delivery", amount: Math.round(c.amount * 100) / 100, needsQuote: !!c.needsQuote,
      });
    }
  }
  const amount = Math.round(lines.reduce((s, l) => s + (l.amount || 0), 0) * 100) / 100;
  return { amount, lines };
}
// Under a cent apart is the same number; anything more is a discrepancy worth showing.
function reconcileInvoice(invoice, items) {
  const expected = computeCategoryTotal(items, {
    client: invoice.client, project: invoice.project, category: invoice.category,
    year: Number(invoice.periodYear), month: Number(invoice.periodMonth),
  });
  const invoiced = Math.round((Number(invoice.amount) || 0) * 100) / 100;
  const difference = Math.round((invoiced - expected.amount) * 100) / 100;
  let status = "match";
  if (!expected.lines.length) status = "nothing";
  else if (Math.abs(difference) >= 0.01) status = difference > 0 ? "over" : "under";
  return { ...expected, invoiced, difference, status };
}
function emptyInvoice() {
  const now = new Date();
  return {
    id: "", number: "", date: todayStr(), dueDate: "", client: CLIENTS[0], project: "",
    orderedBy: "", category: DEBIT_NOTE_CATEGORIES[0],
    periodYear: now.getFullYear(), periodMonth: now.getMonth(),
    amount: "", chargeLine: "", narrative: "", revisedDate: "", revisedBy: "", notes: "", hasFile: false,
  };
}
// Reproduces the debit note as it goes out to the client: same header block, same
// description column with the standard run of lines, same amount in words, same E.&O.E.
// and total row, same signature block. The REVISED stamp only appears when the note has
// actually been revised, matching how it is stamped on the paper ones.
function DebitNotePrint({ invoice, onClose, colors, t }) {
  const money = (n) => `$${(Math.round((Number(n) || 0) * 100) / 100).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  const dmy = (iso) => {
    if (!iso) return "";
    const [y, m, d] = String(iso).split("-");
    return `${Number(d)}/${Number(m)}/${y}`;
  };
  const site = String(invoice.project || "").trim();
  const siteLine = site ? (/^site\s+at/i.test(site) ? site.toUpperCase() : `SITE AT ${site.toUpperCase()}`) : "";
  const addressLines = String(invoice.billTo || "").split("\n").filter(Boolean);
  const cel = { border: "1px solid #111", padding: "6px 10px", verticalAlign: "top" };
  return (
    <div id="debit-note-print-root" className="fixed inset-0 z-50 flex flex-col" style={{ background: "rgba(0,0,0,0.5)" }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #debit-note-print-root, #debit-note-print-root * { visibility: visible; }
          #debit-note-print-root { position: absolute; inset: auto; left: 0; top: 0; width: 100%; background: #fff !important; }
          #debit-note-print-toolbar { display: none !important; }
          #debit-note-print-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          #debit-note-print-area { margin: 0 !important; max-width: 100% !important; }
        }
        @page { size: A4 portrait; margin: 12mm; }
      `}</style>
      <div id="debit-note-print-toolbar" className="flex flex-wrap items-center gap-2 p-3" style={{ background: colors.navy }}>
        <span className="text-sm font-semibold" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>
          {t.debitNotePrintLabel(invoice.number || "")}
        </span>
        <div className="ml-auto flex gap-2">
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => window.print()}>
            {t.printBtn}
          </button>
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.onDark}`, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>
            {t.closePreviewBtn}
          </button>
        </div>
      </div>
      <div id="debit-note-print-scroll" className="flex-1 overflow-y-auto p-6" style={{ background: colors.bg }}>
        <div id="debit-note-print-area" className="mx-auto" style={{ background: "#fff", color: "#111", maxWidth: 720, padding: 28, fontFamily: "Arial, sans-serif", fontSize: 13 }}>

          <div className="flex items-start justify-between">
            <div>
              <div style={{ fontSize: 23, fontWeight: 700, letterSpacing: -0.4 }}>FARSPEED <span style={{ fontWeight: 400 }}>Contractors Limited</span></div>
              <div style={{ fontSize: 11, marginTop: 2 }}>P.O.Box No. 1985, Yuen Long Post Office N.T., Hong Kong</div>
              <div style={{ fontSize: 11 }}>Tel: 5337 9500&nbsp;&nbsp;&nbsp;Fax: 2402 4450&nbsp;&nbsp;&nbsp;www.farspeed.com</div>
            </div>
            <img src={FARSPEED_LOGO_DATA_URI} alt="Farspeed" style={{ height: 66, width: "auto" }} />
          </div>

          <div className="flex items-start justify-between" style={{ marginTop: 26 }}>
            <div style={{ minHeight: 96, paddingTop: 6, fontSize: 13, lineHeight: 1.45 }}>
              <div>{invoice.client}</div>
              {addressLines.map((l, i) => <div key={i}>{l}</div>)}
            </div>
            <div style={{ textAlign: "left" }}>
              {(invoice.revisedDate || invoice.revisedBy) && (
                <div style={{ border: "1px solid #111", padding: "2px 8px", marginBottom: 8, fontSize: 11, width: 168 }}>
                  <div style={{ letterSpacing: 2, textAlign: "center" }}>REVISED</div>
                  <div className="flex justify-between" style={{ borderTop: "1px solid #111", paddingTop: 2 }}>
                    <span>Date: {dmy(invoice.revisedDate)}</span>
                    <span>By: {invoice.revisedBy}</span>
                  </div>
                </div>
              )}
              <div style={{ fontSize: 30, fontWeight: 700 }}>Debit Note</div>
              <div style={{ fontSize: 13, marginTop: 14 }}>No. :{invoice.number}</div>
              <div style={{ fontSize: 13, marginTop: 10 }}>Date: {dmy(invoice.date)}</div>
            </div>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 18 }}>
            <thead>
              <tr>
                <th style={{ ...cel, textAlign: "center", fontWeight: 400 }}>Description</th>
                <th style={{ ...cel, textAlign: "center", fontWeight: 400, width: 150 }}>Amount</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style={{ ...cel, height: 520, lineHeight: 2.1 }}>
                  <div>Your Order Reference:&nbsp; {invoice.orderReference || "STORAGE"}</div>
                  <div>&nbsp;</div>
                  {siteLine && <div>{siteLine}</div>}
                  <div style={{ letterSpacing: -1, lineHeight: 1 }}>--------------------------------------------------------</div>
                  {invoice.orderedBy && <div>ORDERED BY {invoice.orderedBy.toUpperCase()}</div>}
                  {invoice.narrative && <div>{invoice.narrative.toUpperCase()}</div>}
                  {invoice.chargeLine && <div>{invoice.chargeLine.toUpperCase()}</div>}
                  <div>DETAILS AS PER OUR INVENTORY ATTACHED</div>
                  <div style={{ lineHeight: 1.35, marginTop: 8, marginBottom: 8 }}>{amountInWordsHKD(invoice.amount)}</div>
                  {invoice.dueDate && <div>PAYMENT DUE ON {dmy(invoice.dueDate)}</div>}
                </td>
                <td style={{ ...cel, textAlign: "right", paddingTop: 76 }}>{money(invoice.amount)}</td>
              </tr>
              <tr>
                <td style={{ ...cel, textAlign: "center", borderTop: "none" }}>E. &amp; O. E.</td>
                <td style={{ ...cel, borderTop: "none" }}></td>
              </tr>
              <tr>
                <td style={{ ...cel, textAlign: "right" }}>Total Amount: HK$</td>
                <td style={{ ...cel, textAlign: "right" }}>{money(invoice.amount)}</td>
              </tr>
            </tbody>
          </table>

          <div style={{ textAlign: "right", marginTop: 12, fontWeight: 700, fontSize: 15 }}>FARSPEED Contractors Ltd.</div>
          <div style={{ marginTop: 64, marginLeft: "auto", width: 260, borderTop: "1px solid #111" }} />
          <div style={{ textAlign: "center", marginTop: 40, fontSize: 12 }}>a member of FARSPEED Group</div>
        </div>
      </div>
    </div>
  );
}
// The debit notes issued for a month, set against what the depot's own records produce
// for the same client, site, category and month. A note that does not agree is flagged so
// it can be looked at before the client finds it. Storage, CFS, Devan and Delivery only -
// Hoisting and Shifting are a later job.
function InvoicesSection({ items, invoices, setInvoices, monthNames, yearOptions, colors, t, lang }) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth());
  const [client, setClient] = useState("All");
  const [editing, setEditing] = useState(null);
  const [expanded, setExpanded] = useState(null);
  const [printing, setPrinting] = useState(null);
  const [busy, setBusy] = useState("");
  const inputStyle = inputStyleFor(colors);
  const money = (n) => `$${(Math.round((Number(n) || 0) * 100) / 100).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  const inPeriod = (invoices || []).filter((inv) =>
    Number(inv.periodYear) === Number(year) && Number(inv.periodMonth) === Number(month)
    && (client === "All" || inv.client === client));
  const rows = inPeriod.map((inv) => ({ invoice: inv, check: reconcileInvoice(inv, items) }));
  const totalInvoiced = Math.round(rows.reduce((s, r) => s + r.check.invoiced, 0) * 100) / 100;
  const totalExpected = Math.round(rows.reduce((s, r) => s + r.check.amount, 0) * 100) / 100;
  const problems = rows.filter((r) => r.check.status !== "match").length;
  const clientOptions = [...new Set((invoices || []).map((i) => i.client).filter(Boolean))].sort();

  // What the depot produced for this month that no debit note covers yet - the other half
  // of the check, since a charge nobody invoiced is as costly as one invoiced wrongly.
  const uninvoiced = (() => {
    const covered = new Set(inPeriod.map((i) => `${i.client}||${normalizeSiteForMatch(i.project)}||${i.category}`));
    const out = [];
    const sites = new Map();
    for (const item of items || []) {
      if (item.cancelled || !item.client) continue;
      const site = item.constructionSite || item.project || "";
      sites.set(`${item.client}||${site}`, { client: item.client, site });
    }
    for (const { client: c, site } of sites.values()) {
      if (client !== "All" && c !== client) continue;
      for (const category of DEBIT_NOTE_CATEGORIES) {
        const key = `${c}||${normalizeSiteForMatch(site)}||${category}`;
        if (covered.has(key)) continue;
        const got = computeCategoryTotal(items, { client: c, project: site, category, year: Number(year), month: Number(month) });
        if (got.amount > 0) out.push({ client: c, site, category, amount: got.amount });
      }
    }
    return out;
  })();

  function saveInvoice(draft) {
    const id = draft.id || `INV${Date.now()}${Math.floor(Math.random() * 1000)}`;
    setInvoices((prev) => {
      const list = prev || [];
      return list.some((i) => i.id === id)
        ? list.map((i) => (i.id === id ? { ...draft, id } : i))
        : [...list, { ...draft, id }];
    });
    setEditing(null);
  }
  function removeInvoice(id) {
    setInvoices((prev) => (prev || []).filter((i) => i.id !== id));
    storageSet(`invoiceDoc:${id}`, "").catch(() => {});
  }
  async function viewScan(id) {
    setBusy(id);
    try {
      const res = await storageGet(`invoiceDoc:${id}`);
      const doc = res && res.value ? JSON.parse(res.value) : null;
      if (doc && doc.uri) {
        const w = window.open();
        if (w) w.document.write(`<iframe src="${doc.uri}" style="border:0;width:100%;height:100%"></iframe>`);
      }
    } catch (e) { /* nothing stored */ }
    setBusy("");
  }

  const statusStyle = {
    match: { color: colors.green, label: t.invoiceStatusMatch },
    over: { color: colors.red, label: t.invoiceStatusOver },
    under: { color: colors.red, label: t.invoiceStatusUnder },
    nothing: { color: colors.amberText, label: t.invoiceStatusNothing },
  };

  return (
    <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      {printing && <DebitNotePrint invoice={printing} onClose={() => setPrinting(null)} colors={colors} t={t} />}
      <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.invoicesTitle}</h3>
      <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.invoicesDesc}</p>

      <div className="flex flex-wrap items-end gap-3 mb-4">
        <Field label={t.billingMonthLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={month} onChange={(e) => setMonth(Number(e.target.value))}>
            {monthNames.map((m, i) => <option key={m} value={i}>{m}</option>)}
          </select>
        </Field>
        <Field label={t.billingYearLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={year} onChange={(e) => setYear(Number(e.target.value))}>
            {yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </Field>
        <Field label={t.clientLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={client} onChange={(e) => setClient(e.target.value)}>
            <option value="All">{t.statusAll}</option>
            {clientOptions.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <button className="px-3 py-2 rounded text-sm font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
          onClick={() => setEditing({ ...emptyInvoice(), periodYear: year, periodMonth: month, client: client === "All" ? CLIENTS[0] : client })}>
          {t.invoiceAddBtn}
        </button>
      </div>

      {editing && (
        <InvoiceEditor
          draft={editing} onChange={setEditing} onSave={saveInvoice} onCancel={() => setEditing(null)}
          items={items} monthNames={monthNames} yearOptions={yearOptions} colors={colors} t={t}
        />
      )}

      <div className="rounded overflow-hidden mb-3" style={{ border: `1px solid ${colors.line}` }}>
        <table className="w-full text-xs">
          <thead>
            <tr style={{ background: colors.surfaceDim }}>
              {[t.invoiceColNo, t.invoiceColDate, t.clientLabel, t.invoiceColSite, t.invoiceColCategory,
                t.invoiceColInvoiced, t.invoiceColExpected, t.invoiceColDifference, ""].map((h, i) => (
                <th key={i} className="text-left px-3 py-2 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr><td colSpan={10} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.invoicesNoneMsg}</td></tr>
            )}
            {rows.map(({ invoice, check }) => {
              const st = statusStyle[check.status];
              const open = expanded === invoice.id;
              return (
                <React.Fragment key={invoice.id}>
                  <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }}
                    onClick={() => setExpanded(open ? null : invoice.id)}>
                    <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{invoice.number || "—"}</td>
                    <td className="px-3 py-2">{invoice.date ? fmt(invoice.date) : "—"}</td>
                    <td className="px-3 py-2">{invoice.client}</td>
                    <td className="px-3 py-2 max-w-[200px] truncate">{invoice.project || "—"}</td>
                    <td className="px-3 py-2">{invoice.category}</td>
                    <td className="px-3 py-2 text-right" style={{ fontFamily: FONT_MONO }}>{money(check.invoiced)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontFamily: FONT_MONO }}>{money(check.amount)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontFamily: FONT_MONO, color: st.color, fontWeight: 600 }}>
                      {check.status === "match" ? st.label : `${check.difference > 0 ? "+" : ""}${money(check.difference)}`}
                    </td>
                    <td className="px-3 py-2 text-right whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                      {invoice.hasFile && (
                        <button className="text-xs font-semibold mr-2" style={{ color: colors.amberText }} onClick={() => viewScan(invoice.id)}>
                          {busy === invoice.id ? "…" : t.invoiceViewScanBtn}
                        </button>
                      )}
                      <button className="text-xs font-semibold mr-2" style={{ color: colors.amberText }} onClick={() => setPrinting(invoice)}>{t.printBtn}</button>
                      <button className="text-xs font-semibold mr-2" style={{ color: colors.inkFaint }} onClick={() => setEditing(invoice)}>{t.editBtn}</button>
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => removeInvoice(invoice.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                  {open && (
                    <tr style={{ background: colors.surfaceDim }}>
                      <td colSpan={9} className="px-3 py-2">
                        <div className="text-xs font-semibold mb-1" style={{ color: st.color }}>
                          {check.status === "nothing" ? t.invoiceNothingHint : `${st.label} \u00b7 ${t.invoiceLinesLabel(check.lines.length)}`}
                        </div>
                        {check.lines.map((l, i) => (
                          <div key={i} className="text-xs flex gap-3" style={{ color: colors.ink }}>
                            <span style={{ fontFamily: FONT_MONO, minWidth: 70 }}>{l.itemId}</span>
                            <span style={{ minWidth: 90 }}>{l.unitCode || l.jobNumber || ""}</span>
                            <span className="flex-1">{l.detail}{l.estimated ? ` \u00b7 ${t.invoiceEstimatedTag}` : ""}</span>
                            <span style={{ fontFamily: FONT_MONO }}>{money(l.amount)}</span>
                          </div>
                        ))}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="flex flex-wrap gap-4 text-sm mb-2" style={{ color: colors.ink }}>
        <span>{t.invoiceTotalInvoiced}: <strong style={{ fontFamily: FONT_MONO }}>{money(totalInvoiced)}</strong></span>
        <span>{t.invoiceTotalExpected}: <strong style={{ fontFamily: FONT_MONO }}>{money(totalExpected)}</strong></span>
        <span style={{ color: Math.abs(totalInvoiced - totalExpected) >= 0.01 ? colors.red : colors.green }}>
          {t.invoiceTotalDifference}: <strong style={{ fontFamily: FONT_MONO }}>{money(totalInvoiced - totalExpected)}</strong>
        </span>
        {problems > 0 && <span style={{ color: colors.red }}>{t.invoiceProblemCount(problems)}</span>}
      </div>

      {uninvoiced.length > 0 && (
        <div className="rounded p-3 mt-3" style={{ background: colors.redSoft }}>
          <div className="text-xs font-semibold mb-1" style={{ color: colors.red }}>{t.invoiceUninvoicedLabel}</div>
          {uninvoiced.map((u, i) => (
            <div key={i} className="text-xs flex gap-3" style={{ color: colors.ink }}>
              <span style={{ minWidth: 110 }}>{u.client}</span>
              <span className="flex-1 truncate">{u.site}</span>
              <span style={{ minWidth: 70 }}>{u.category}</span>
              <span style={{ fontFamily: FONT_MONO }}>{money(u.amount)}</span>
            </div>
          ))}
        </div>
      )}
      <div className="text-xs mt-3" style={{ color: colors.inkFaint }}>{t.invoicesFootnote}</div>
    </div>
  );
}
// The debit note's own fields. The site, category and month decide which of the depot's
// charges the note is checked against, so those three matter more than the rest.
function InvoiceEditor({ draft, onChange, onSave, onCancel, items, monthNames, yearOptions, colors, t }) {
  const inputStyle = inputStyleFor(colors);
  const [attaching, setAttaching] = useState(false);
  const set = (k) => (e) => onChange({ ...draft, [k]: e.target.value });
  const siteOptions = [...new Set((items || []).map((i) => i.constructionSite || i.project).filter(Boolean))].sort();
  const preview = reconcileInvoice(draft, items);
  const money = (n) => `$${(Math.round((Number(n) || 0) * 100) / 100).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  async function attach(file) {
    if (!file) return;
    setAttaching(true);
    const id = draft.id || `INV${Date.now()}${Math.floor(Math.random() * 1000)}`;
    try {
      const uri = await compressFileToDataUri(file, 1600, 0.7);
      await storageSet(`invoiceDoc:${id}`, JSON.stringify({ uri, name: file.name, at: todayStr() }));
      onChange({ ...draft, id, hasFile: true, scanName: file.name });
    } catch (e) { /* leave unattached */ }
    setAttaching(false);
  }

  return (
    <div className="rounded p-3 mb-4" style={{ border: `1px dashed ${colors.line}`, background: colors.amberSoft }}>
      <div className="flex flex-wrap gap-3 mb-2">
        <Field label={t.invoiceColNo} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 130 }} value={draft.number} onChange={set("number")} placeholder="02608003" />
        </Field>
        <Field label={t.invoiceColDate} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={draft.date} onChange={set("date")} />
        </Field>
        <Field label={t.invoiceDueDateLabel} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={draft.dueDate} onChange={set("dueDate")} />
        </Field>
        <Field label={t.clientLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={draft.client} onChange={set("client")}>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.invoiceColCategory} colors={colors}>
          <select className={inputClass} style={inputStyle} value={draft.category} onChange={set("category")}>
            {DEBIT_NOTE_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.billingMonthLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={draft.periodMonth} onChange={(e) => onChange({ ...draft, periodMonth: Number(e.target.value) })}>
            {monthNames.map((m, i) => <option key={m} value={i}>{m}</option>)}
          </select>
        </Field>
        <Field label={t.billingYearLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={draft.periodYear} onChange={(e) => onChange({ ...draft, periodYear: Number(e.target.value) })}>
            {yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </Field>
        <Field label={t.invoiceAmountLabel} colors={colors}>
          <input type="number" min="0" step="0.01" className={inputClass} style={{ ...inputStyle, width: 130 }} value={draft.amount} onChange={set("amount")} />
        </Field>
      </div>
      <div className="flex flex-wrap gap-3 mb-2">
        <Field label={t.invoiceColSite} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 360 }} value={draft.project} onChange={set("project")} list="invoice-site-options" />
          <datalist id="invoice-site-options">{siteOptions.map((s) => <option key={s} value={s} />)}</datalist>
        </Field>
        <Field label={t.invoiceOrderedByLabel} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 200 }} value={draft.orderedBy} onChange={set("orderedBy")} />
        </Field>
        <Field label={t.invoiceOrderRefLabel} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 140 }} value={draft.orderReference || ""} onChange={set("orderReference")} placeholder="STORAGE" />
        </Field>
      </div>
      <div className="flex flex-wrap gap-3 mb-2">
        <Field label={t.invoiceChargeLineLabel} hint={t.invoiceChargeLineHint} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 460 }} value={draft.chargeLine} onChange={set("chargeLine")}
            placeholder="CHARGES FOR OPEN YARD STORAGE AREA AUG 2026 AT LUMPSUM" />
        </Field>
        <Field label={t.invoiceNarrativeLabel} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 380 }} value={draft.narrative} onChange={set("narrative")}
            placeholder="COVERED AREA STORAGE AT FARSPEED DEPOT AS INSTRUCTED." />
        </Field>
      </div>
      <div className="flex flex-wrap gap-3 mb-2">
        <Field label={t.invoiceBillToLabel} colors={colors}>
          <textarea className={inputClass} style={{ ...inputStyle, width: 340, height: 70 }} value={draft.billTo || ""} onChange={set("billTo")} />
        </Field>
        <Field label={t.invoiceRevisedDateLabel} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={draft.revisedDate} onChange={set("revisedDate")} />
        </Field>
        <Field label={t.invoiceRevisedByLabel} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, width: 120 }} value={draft.revisedBy} onChange={set("revisedBy")} />
        </Field>
        <Field label={t.invoiceScanLabel} colors={colors}>
          <input type="file" accept="image/*,application/pdf" className="text-xs" onChange={(e) => attach(e.target.files && e.target.files[0])} />
        </Field>
      </div>
      <div className="text-xs mb-2" style={{ color: preview.status === "match" ? colors.green : colors.red }}>
        {attaching ? t.invoiceAttachingMsg : t.invoicePreviewMsg(money(preview.amount), money(preview.difference))}
      </div>
      <div className="flex gap-2">
        <button className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
          onClick={() => onSave(draft)}>{t.saveBtn}</button>
        <button className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={onCancel}>{t.cancelBtn}</button>
      </div>
    </div>
  );
}
function BillingPanel({ items, invoices, setInvoices, onDeleteItem, onDeleteItems, authUser, colors, t, lang }) {
  const now = new Date();
  const [mode, setMode] = useState("search");
  const [search, setSearch] = useState("");
  const [filterClient, setFilterClient] = useState("All");
  const [filterProject, setFilterProject] = useState("All");
  const [filterJobNo, setFilterJobNo] = useState("All");
  const [expanded, setExpanded] = useState(null);
  const [summaryYear, setSummaryYear] = useState(now.getFullYear());
  const [summaryMonth, setSummaryMonth] = useState(now.getMonth());
  const [expandedClient, setExpandedClient] = useState(null);
  const [pendingDeleteItem, setPendingDeleteItem] = useState(null);
  // Selection is by entry, not by row: one entry can produce several billing rows (one per
  // arrival batch), and deleting removes the entry behind all of them. Ticking any of its
  // rows ticks them all, so what gets deleted is never a surprise.
  const [selectedForDelete, setSelectedForDelete] = useState([]);
  const [pendingDeleteMany, setPendingDeleteMany] = useState(false);

  const [expandedHandlingClient, setExpandedHandlingClient] = useState(null);
  const [expandedHandlingMonth, setExpandedHandlingMonth] = useState(null);

  const allRows = useMemo(() => {
    const rows = [];
    for (const item of items) {
      for (const r of computeItemBillingRows(item)) rows.push(r);
    }
    return rows;
  }, [items]);

  const clientOptions = useMemo(() => [...new Set(allRows.map((r) => r.item.client).filter(Boolean))].sort(), [allRows]);
  const projectOptions = useMemo(() => [...new Set(allRows.map((r) => r.item.constructionSite || r.item.project).filter(Boolean))].sort(), [allRows]);
  const jobNoOptions = useMemo(() => [...new Set(allRows.map((r) => r.item.jobNumber).filter(Boolean))].sort(), [allRows]);

  const monthlySummary = useMemo(() => computeMonthlyBillingSummary(items, summaryYear, summaryMonth), [items, summaryYear, summaryMonth]);
  const yearOptions = useMemo(() => {
    const ys = new Set([now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1]);
    for (const item of items) {
      const d = item.depotArrivalDate || effectiveDepotArrivalDate(item);
      if (d) ys.add(toDateOnly(d).getFullYear());
    }
    return [...ys].sort();
  }, [items]);
  const monthNames = lang === "zh"
    ? ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
    : ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const filtered = useMemo(() => {
    return allRows.filter((r) => {
      const i = r.item;
      if (filterClient !== "All" && i.client !== filterClient) return false;
      if (filterProject !== "All" && (i.constructionSite || i.project) !== filterProject) return false;
      if (filterJobNo !== "All" && i.jobNumber !== filterJobNo) return false;
      if (!search.trim()) return true;
      const q = search.toLowerCase();
      return (
        i.client?.toLowerCase().includes(q) || i.project?.toLowerCase().includes(q) ||
        i.constructionSite?.toLowerCase().includes(q) || i.jobNumber?.toLowerCase().includes(q) ||
        i.id?.toLowerCase().includes(q) || i.shkNumber?.toLowerCase().includes(q) ||
        (r.codes || []).some((c) => (c || "").toLowerCase().includes(q))
      );
    });
  }, [allRows, search, filterClient, filterProject, filterJobNo]);

  const grandTotal = Math.round(filtered.reduce((s, r) => s + r.total, 0) * 100) / 100;
  // The entries currently on screen, and what ticking them would remove: an entry may
  // account for several rows, so both counts are shown before anyone confirms.
  const visibleDeleteIds = [...new Set(filtered.map((r) => r.item.id))];
  const selectedRows = filtered.filter((r) => selectedForDelete.includes(r.item.id));
  const selectedRowCount = selectedRows.length;
  const selectedRowTotal = Math.round(selectedRows.reduce((s, r) => s + r.total, 0) * 100) / 100;

  const [handlingYear, setHandlingYear] = useState(now.getFullYear());
  const handlingRows = useMemo(() => {
    const rows = [];
    for (const item of items) {
      if (!HANDLING_TARIFFS[item.client]) continue;
      const hasArrival = (item.arrivals || []).length > 0 || item.depotArrivalDate;
      const hasDelivery = (item.deliveries || []).length > 0;
      const arrivalDate = effectiveDepotArrivalDate(item);
      if (hasArrival) {
        const c = computeHandlingCharge(item, "devan");
        if (c) rows.push({ item, jobType: "devan", date: arrivalDate, ...c });
      }
      if (hasDelivery) {
        const c = computeHandlingCharge(item, "delivery");
        if (c) rows.push({ item, jobType: "delivery", date: lastDeliveryDate(item), ...c });
      }
      const haul = computeContainerHaulageCharge(item);
      if (haul) rows.push({ item, jobType: "haulage", date: arrivalDate, ...haul });
    }
    return rows;
  }, [items]);
  const handlingGrandTotal = Math.round(handlingRows.reduce((s, r) => s + (r.amount || 0), 0) * 100) / 100;
  const handlingNeedsQuoteCount = handlingRows.filter((r) => r.needsQuote).length;
  const handlingYears = useMemo(() => {
    const ys = new Set(handlingRows.filter((r) => r.date).map((r) => Number(r.date.slice(0, 4))));
    ys.add(now.getFullYear());
    return [...ys].sort((a, b) => b - a);
  }, [handlingRows]);
  // Groups this year's rows by client, then by month, with a yearly subtotal per client.
  const handlingGrouped = useMemo(() => {
    const byClient = new Map();
    for (const r of handlingRows) {
      if (!r.date || Number(r.date.slice(0, 4)) !== handlingYear) continue;
      if (!byClient.has(r.item.client)) byClient.set(r.item.client, { client: r.item.client, yearTotal: 0, months: new Map() });
      const g = byClient.get(r.item.client);
      const monthIdx = Number(r.date.slice(5, 7)) - 1;
      if (!g.months.has(monthIdx)) g.months.set(monthIdx, { monthIdx, total: 0, rows: [] });
      const m = g.months.get(monthIdx);
      m.rows.push(r);
      if (r.amount) { m.total += r.amount; g.yearTotal += r.amount; }
    }
    return [...byClient.values()].map((g) => ({
      ...g,
      months: [...g.months.values()].sort((a, b) => b.monthIdx - a.monthIdx),
    })).sort((a, b) => a.client.localeCompare(b.client));
  }, [handlingRows, handlingYear]);
  const money = (n) => `$${n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  return (
    <div className="flex flex-col gap-4">
      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.billingTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.billingDesc}</p>
        <div className="flex gap-1 rounded-lg p-1 mb-3" style={{ background: colors.surfaceDim, width: "fit-content" }}>
          {[["search", t.billingModeSearch], ["monthly", t.billingModeMonthly], ["handling", t.billingModeHandling], ["invoices", t.billingModeInvoices]].map(([k, label]) => (
            <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
              {label}
            </button>
          ))}
        </div>
        {mode === "search" && (
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.searchLabel} colors={colors}>
            <input
              className={inputClass}
              style={{ ...inputStyleFor(colors), minWidth: 240 }}
              placeholder={t.billingSearchPlaceholder}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </Field>
          <Field label={t.clientLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={filterClient} onChange={(e) => setFilterClient(e.target.value)}>
              <option>All</option>
              {clientOptions.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label={t.billingColProject} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={filterProject} onChange={(e) => setFilterProject(e.target.value)}>
              <option>All</option>
              {projectOptions.map((p) => <option key={p}>{p}</option>)}
            </select>
          </Field>
          <Field label={t.billingColJobNo} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={filterJobNo} onChange={(e) => setFilterJobNo(e.target.value)}>
              <option>All</option>
              {jobNoOptions.map((j) => <option key={j}>{j}</option>)}
            </select>
          </Field>
        </div>
        )}
        {mode === "monthly" && (
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.billingMonthLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={summaryMonth} onChange={(e) => setSummaryMonth(Number(e.target.value))}>
              {monthNames.map((m, idx) => <option key={idx} value={idx}>{m}</option>)}
            </select>
          </Field>
          <Field label={t.billingYearLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={summaryYear} onChange={(e) => setSummaryYear(Number(e.target.value))}>
              {yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
          </Field>
        </div>
        )}
      </div>

      {mode === "monthly" && (
        <>
          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.billingColClient, t.billingColTotal, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {monthlySummary.clients.length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.billingMonthNoneMsg}</td></tr>
                )}
                {monthlySummary.clients.map((g) => {
                  const isOpen = expandedClient === g.client;
                  return (
                    <React.Fragment key={g.client}>
                      <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }} onClick={() => setExpandedClient(isOpen ? null : g.client)}>
                        <td className="px-3 py-2 font-semibold">{g.client}</td>
                        <td className="px-3 py-2 font-semibold">{money(g.total)}</td>
                        <td className="px-3 py-2 text-right text-xs" style={{ color: colors.amberText }}>{isOpen ? t.billingHideBtn : t.billingShowBtn}</td>
                      </tr>
                      {isOpen && (
                        <tr style={{ background: colors.surfaceDim }}>
                          <td colSpan={3} className="px-4 py-3">
                            <table className="w-full text-xs" style={{ color: colors.ink }}>
                              <thead>
                                <tr>
                                  {[t.billingColProject, t.billingColJobNo, t.billingColCbm, t.billingColBatchDate, "", ""].map((h) => (
                                    <th key={h} className="text-left pr-4 pb-1 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {g.lines.map((l, i) => (
                                  <tr key={i} style={{ borderTop: `1px solid ${colors.line}` }}>
                                    <td className="pr-4 py-1">{l.project}</td>
                                    <td className="pr-4 py-1" style={{ fontFamily: FONT_MONO }}>{l.jobNumber || "—"}</td>
                                    <td className="pr-4 py-1">{l.cbm.toFixed(3)}{l.estimated ? " *" : ""}</td>
                                    <td className="pr-4 py-1">{fmt(l.batchDate)}</td>
                                    <td className="pr-4 py-1" style={{ color: colors.inkFaint }}>{l.detail}</td>
                                    <td className="py-1 text-right font-semibold">{money(l.amount)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
              {monthlySummary.clients.length > 0 && (
                <tfoot>
                  <tr style={{ borderTop: `2px solid ${colors.line}` }}>
                    <td className="px-3 py-2 text-right font-semibold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.billingGrandTotal}</td>
                    <td className="px-3 py-2 font-bold" style={{ color: colors.ink }}>{money(monthlySummary.grandTotal)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
          <div className="text-xs" style={{ color: colors.inkFaint }}>{t.billingMonthFootnote}</div>
        </>
      )}

      {mode === "search" && (
      <>
        {selectedForDelete.length > 0 && (
          <div className="rounded-lg px-4 py-3 mb-3 flex flex-wrap items-center gap-3" style={{ background: colors.redSoft, border: `1px solid ${colors.red}` }}>
            <span className="text-sm font-semibold" style={{ color: colors.red }}>
              {t.billingSelectedCount(selectedForDelete.length, selectedRowCount, money(selectedRowTotal))}
            </span>
            <button
              className="px-3 py-1.5 rounded text-xs font-semibold"
              style={{ background: colors.red, color: colors.onDark, fontFamily: FONT_DISPLAY }}
              onClick={() => setPendingDeleteMany(true)}
            >
              {t.billingDeleteSelectedBtn(selectedForDelete.length)}
            </button>
            <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => setSelectedForDelete([])}>
              {t.clearBtn}
            </button>
            <span className="text-xs" style={{ color: colors.red }}>{t.billingDeleteSelectedHint}</span>
          </div>
        )}
        <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
        <table className="w-full text-sm" style={{ background: colors.surface }}>
          <thead>
            <tr style={{ background: colors.surfaceDim }}>
              <th className="px-3 py-2 w-8">
                <input
                  type="checkbox"
                  aria-label={t.billingSelectAllAria}
                  checked={visibleDeleteIds.length > 0 && visibleDeleteIds.every((id) => selectedForDelete.includes(id))}
                  onChange={(e) => setSelectedForDelete(e.target.checked
                    ? [...new Set([...selectedForDelete, ...visibleDeleteIds])]
                    : selectedForDelete.filter((id) => !visibleDeleteIds.includes(id)))}
                />
              </th>
              {[t.billingColClient, t.billingColProject, t.billingColJobNo, t.billingColBatchDate, t.billingColCbm, t.billingColRate, t.billingColStatus, t.billingColTotal, "", ""].map((h, idx) => (
                <th key={idx} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={11} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.billingNoneMsg}</td></tr>
            )}
            {filtered.map((r, idx) => {
              const key = `${r.item.id}-${idx}`;
              const isOpen = expanded === key;
              return (
                <React.Fragment key={key}>
                  <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }} onClick={() => setExpanded(isOpen ? null : key)}>
                    <td className="px-3 py-2" onClick={(e) => e.stopPropagation()}>
                      <input
                        type="checkbox"
                        checked={selectedForDelete.includes(r.item.id)}
                        onChange={(e) => setSelectedForDelete((prev) => (e.target.checked
                          ? [...new Set([...prev, r.item.id])]
                          : prev.filter((id) => id !== r.item.id)))}
                      />
                    </td>
                    <td className="px-3 py-2">{r.item.client}</td>
                    <td className="px-3 py-2">{r.item.constructionSite || r.item.project}</td>
                    <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.item.jobNumber || "—"}</td>
                    <td className="px-3 py-2">{fmt(r.batchDate)}</td>
                    <td className="px-3 py-2">{r.cbm.toFixed(3)}{r.estimated ? <span title={t.billingEstimatedNote} style={{ color: colors.amberText, cursor: "help" }}> *</span> : ""}</td>
                    <td className="px-3 py-2">${r.rate}/{t.billingPerCbmMonth}</td>
                    <td className="px-3 py-2">
                      {r.ongoing ? <Badge tone="amber" colors={colors}>{t.billingOngoing}</Badge> : <Badge tone="grey" colors={colors}>{t.billingClosed}</Badge>}
                    </td>
                    <td className="px-3 py-2 font-semibold">{money(r.total)}</td>
                    <td className="px-3 py-2 text-right text-xs" style={{ color: colors.amberText }}>{isOpen ? t.billingHideBtn : t.billingShowBtn}</td>
                    <td className="px-3 py-2 text-right">
                      <button
                        type="button"
                        title={t.billingDeleteItemBtn}
                        aria-label={t.billingDeleteItemBtn}
                        className="w-6 h-6 rounded-full inline-flex items-center justify-center font-bold"
                        style={{ background: colors.redSoft, color: colors.red, lineHeight: 1 }}
                        onClick={(e) => { e.stopPropagation(); setPendingDeleteItem(r.item); }}
                      >
                        &minus;
                      </button>
                    </td>
                  </tr>
                  {isOpen && (
                    <tr style={{ background: colors.surfaceDim }}>
                      <td colSpan={11} className="px-4 py-3">
                        <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>
                          {t.billingFreeDaysNote(r.freeDays)}{r.codes && r.codes.length > 0 ? ` · ${t.billingCasesLabel}: ${r.codes.join(", ")}` : ""}
                        </div>
                        <table className="text-xs" style={{ color: colors.ink }}>
                          <tbody>
                            {r.breakdown.map((b, i) => (
                              <tr key={i}>
                                <td className="pr-4 py-0.5 font-semibold">{b.label}</td>
                                <td className="pr-4 py-0.5" style={{ color: colors.inkFaint }}>{b.detail}</td>
                                <td className="py-0.5 text-right">{money(b.amount)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        {r.estimated && <div className="text-xs mt-2" style={{ color: colors.amberText }}>{t.billingEstimatedNote}</div>}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
          {filtered.length > 0 && (
            <tfoot>
              <tr style={{ borderTop: `2px solid ${colors.line}` }}>
                <td colSpan={7} className="px-3 py-2 text-right font-semibold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.billingGrandTotal}</td>
                <td className="px-3 py-2 font-bold" style={{ color: colors.ink }}>{money(grandTotal)}</td>
                <td></td>
                <td></td>
              </tr>
            </tfoot>
          )}
        </table>
        </div>
        <div className="text-xs" style={{ color: colors.inkFaint }}>{t.billingFootnote}</div>
      </>
      )}
      {mode === "invoices" && (
        <InvoicesSection
          items={items} invoices={invoices} setInvoices={setInvoices}
          monthNames={monthNames} yearOptions={yearOptions} colors={colors} t={t} lang={lang}
        />
      )}
      {mode === "handling" && (
        <>
          <div className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.billingHandlingDesc}</div>
          {handlingNeedsQuoteCount > 0 && (
            <div className="px-3 py-2 rounded text-sm mb-3" style={{ background: colors.amberSoft, color: colors.amberText }}>
              {t.billingHandlingNeedsQuote(handlingNeedsQuoteCount)}
            </div>
          )}
          <div className="flex items-end gap-3 mb-3">
            <Field label={t.billingYearLabel} colors={colors}>
              <select className={inputClass} style={inputStyleFor(colors)} value={handlingYear} onChange={(e) => { setHandlingYear(Number(e.target.value)); setExpandedHandlingClient(null); setExpandedHandlingMonth(null); }}>
                {handlingYears.map((y) => <option key={y} value={y}>{y}</option>)}
              </select>
            </Field>
          </div>
          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.billingColClient, t.billingColTotal, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {handlingGrouped.length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.billingHandlingNoneMsg}</td></tr>
                )}
                {handlingGrouped.map((g) => {
                  const isClientOpen = expandedHandlingClient === g.client;
                  return (
                    <React.Fragment key={g.client}>
                      <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }} onClick={() => { setExpandedHandlingClient(isClientOpen ? null : g.client); setExpandedHandlingMonth(null); }}>
                        <td className="px-3 py-2 font-semibold">{g.client}</td>
                        <td className="px-3 py-2 font-semibold">{money(g.yearTotal)}</td>
                        <td className="px-3 py-2 text-right text-xs" style={{ color: colors.amberText }}>{isClientOpen ? t.billingHideBtn : t.billingShowBtn}</td>
                      </tr>
                      {isClientOpen && g.months.map((m) => {
                        const monthKey = `${g.client}-${m.monthIdx}`;
                        const isMonthOpen = expandedHandlingMonth === monthKey;
                        return (
                          <React.Fragment key={monthKey}>
                            <tr style={{ background: colors.surfaceDim, cursor: "pointer" }} onClick={() => setExpandedHandlingMonth(isMonthOpen ? null : monthKey)}>
                              <td className="px-4 py-2 pl-8" style={{ color: colors.ink }}>{monthNames[m.monthIdx]}</td>
                              <td className="px-4 py-2 font-semibold" style={{ color: colors.ink }}>{money(m.total)}</td>
                              <td className="px-3 py-2 text-right text-xs" style={{ color: colors.amberText }}>{isMonthOpen ? t.billingHideBtn : t.billingShowBtn}</td>
                            </tr>
                            {isMonthOpen && (
                              <tr>
                                <td colSpan={3} className="px-4 py-3" style={{ background: colors.surfaceDim }}>
                                  <table className="w-full text-xs" style={{ color: colors.ink }}>
                                    <thead>
                                      <tr>
                                        {[t.billingColProject, t.billingColJobNo, t.billingHandlingColType, t.billingHandlingColBasis, t.billingHandlingColRate, ""].map((h) => (
                                          <th key={h} className="text-left pr-4 pb-1 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                                        ))}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {m.rows.map((r, idx) => (
                                        <tr key={idx} style={{ borderTop: `1px solid ${colors.line}` }}>
                                          <td className="pr-4 py-1">{r.item.project || r.item.constructionSite || "—"}</td>
                                          <td className="pr-4 py-1" style={{ fontFamily: FONT_MONO }}>{r.item.jobNumber || "—"}</td>
                                          <td className="pr-4 py-1">
                                            {r.jobType === "devan" ? t.billingHandlingTypeDevan : r.jobType === "delivery" ? t.billingHandlingTypeDelivery : t.billingHandlingTypeHaulage}
                                            {r.oversizeMult > 1 && <span className="ml-1" style={{ color: colors.amberText }}>{t.billingHandlingOversizeTag(r.oversizeMult)}</span>}
                                          </td>
                                          <td className="pr-4 py-1" style={{ color: colors.inkFaint }}>
                                            {r.jobType === "haulage"
                                              ? t.billingHandlingHaulageBasis(r.containers20, r.containers40)
                                              : `${Math.round((r.rTons || 0) * 1000) / 1000} R/Ton \u00b7 ${r.zone === "lantau" ? t.zoneLantau : t.zoneUrban} \u00b7 ${r.cargoType === "escalator" ? t.cargoTypeEscalator : t.cargoTypeElevator}`}
                                          </td>
                                          <td className="pr-4 py-1" style={{ color: colors.inkFaint }}>{r.jobType !== "haulage" && r.perRTonRate ? `$${r.perRTonRate}/R.Ton` : "—"}</td>
                                          <td className="py-1 text-right font-semibold">{r.needsQuote ? <span style={{ color: colors.amberText }}>{t.billingHandlingQuoteBadge}</span> : money(r.amount)}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </React.Fragment>
                  );
                })}
              </tbody>
              {handlingGrouped.length > 0 && (
                <tfoot>
                  <tr style={{ borderTop: `2px solid ${colors.line}` }}>
                    <td className="px-3 py-2 text-right font-semibold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.billingGrandTotal}</td>
                    <td className="px-3 py-2 font-bold" style={{ color: colors.ink }}>{money(Math.round(handlingGrouped.reduce((s, g) => s + g.yearTotal, 0) * 100) / 100)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
          <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>{t.billingHandlingFootnote}</div>
        </>
      )}
      {pendingDeleteMany && (
        <AdminConfirmModal
          authUser={authUser}
          onConfirm={() => {
            // One password for the batch, and one write: see handleDeleteMany.
            if (onDeleteItems) onDeleteItems(selectedForDelete);
            else selectedForDelete.forEach((id) => onDeleteItem(id));
            setSelectedForDelete([]);
            setPendingDeleteMany(false);
          }}
          onClose={() => setPendingDeleteMany(false)}
          colors={colors}
          t={t}
        />
      )}
      {pendingDeleteItem && (
        <AdminConfirmModal
          authUser={authUser}
          onConfirm={() => { onDeleteItem(pendingDeleteItem.id); setPendingDeleteItem(null); }}
          onClose={() => setPendingDeleteItem(null)}
          colors={colors}
          t={t}
        />
      )}
    </div>
  );
}

const LEGACY_DOC_TYPES = JOB_SHEET_TEMPLATES;
function guessDocTypeFromName(name) {
  const n = (name || "").toLowerCase();
  const map = [
    ["devan", "Devan"], ["cfs", "CFS"], ["delivery", "Delivery"], ["return", "Return"],
    ["shifting", "Shifting"], ["hoisting", "Hoisting"], ["day_work", "Day Work"], ["day work", "Day Work"],
    ["dismantle", "Dismantle"], ["dis_n_removal", "Dis & Removal of Lifting Tools"], ["removal", "Dis & Removal of Lifting Tools"],
    ["job_cancel", "Job Cancel"], ["cancel", "Job Cancel"], ["pick-up", "Pick-up"], ["pickup", "Pick-up"],
    ["re-position", "Re-position"], ["reposition", "Re-position"], ["position", "Position"],
    ["retain", "Retain of Safety Ropes"],
  ];
  for (const [needle, type] of map) if (n.includes(needle)) return type;
  return "Devan";
}
// A staged row usually comes from one uploaded file, so its name is the file's. A row that
// came out of the job sheet importer's spreadsheet is one of many carried in a single
// upload, and what belongs on its arrival and in the archive is the job sheet it was read
// from - "Devan_1_2606087.xlsx" - not the spreadsheet that ferried a hundred of them in.
function legacySourceName(row) {
  return (row && (row.sourceName || (row.file && row.file.name))) || "";
}
// The columns the job sheet importer writes. Recognised by the ones that carry meaning
// rather than by all of them, so a spreadsheet someone has added a column to still reads.
const JOBSHEET_SS_REQUIRED = ["file name", "job number", "case numbers", "lot / dm no."];
function jobSheetSpreadsheetColumns(headerRow) {
  const map = {};
  (headerRow || []).forEach((cell, i) => {
    const key = String(cell == null ? "" : cell).trim().toLowerCase();
    if (key && map[key] === undefined) map[key] = i;
  });
  return JOBSHEET_SS_REQUIRED.every((k) => map[k] !== undefined) ? map : null;
}
// Turns that spreadsheet back into the same staged rows an uploaded job sheet produces, so
// everything downstream - matching to a shipment, pre-selecting cases, reporting what is not
// at the depot, letting someone correct it by hand - is the machinery that already exists
// rather than a second version of it.
//
// One job sheet becomes one staged row however many lots it had, because that is what it was
// on the way out: the spreadsheet's rows are lots, and the File Name column says which sheet
// they came from. Rows are grouped back by that name, and by job number under it, so two
// sheets that somehow share a file name don't merge.
function rowsFromJobSheetSpreadsheet(grid, file, resolveClient) {
  const headerIdx = (grid || []).findIndex((r) => jobSheetSpreadsheetColumns(r));
  if (headerIdx === -1) return null;
  const col = jobSheetSpreadsheetColumns(grid[headerIdx]);
  const cell = (row, name) => {
    const i = col[name];
    return i === undefined ? "" : String((row || [])[i] == null ? "" : (row || [])[i]).trim();
  };
  const num = (v) => {
    const n = Number(String(v || "").replace(/[, ]/g, ""));
    return isFinite(n) && n !== 0 ? n : "";
  };

  const bySheet = new Map();
  for (let i = headerIdx + 1; i < grid.length; i++) {
    const r = grid[i] || [];
    if (r.every((c) => String(c == null ? "" : c).trim() === "")) continue;
    const sourceName = cell(r, "file name") || cell(r, "job number") || "row " + (i + 1);
    const key = `${sourceName}\u0000${cell(r, "job number")}`;
    if (!bySheet.has(key)) bySheet.set(key, { sourceName, lots: [] });
    bySheet.get(key).lots.push(r);
  }
  if (!bySheet.size) return null;

  const out = [];
  for (const sheet of bySheet.values()) {
    const first = sheet.lots[0];
    const refBlocks = [];
    const caseCodesByLot = {};
    const caseMarksByLot = {};
    const declaredTotalsList = [];
    const caseCountMismatches = [];
    let pkgs = 0, kg = 0, cbm = 0;

    for (const r of sheet.lots) {
      const codes = cell(r, "case numbers").split(",").map((c) => c.trim()).filter(Boolean);
      // Numbers and markings are held apart exactly as a read job sheet holds them, so the
      // same matching runs: a lift's plain "4, 5, 6" is a case mark, "09B1109" is a code.
      // A padded list is a marking too - an escalator's "02, 061, 081" is not two, sixty-one
      // and eighty-one - and turning it back into numbers on the way in would undo on the
      // return leg exactly what the sheet reader was careful to keep.
      const allNumeric = codes.length > 0
        && codes.every((c) => /^\d+$/.test(c))
        && !codes.some((c) => /^0\d/.test(c));
      const lot = {
        lotRef: cell(r, "lot / dm no.") || cell(r, "lift no."),
        altRef: "", unitCode: cell(r, "lift no."),
        caseNumbers: allNumeric ? codes.map((c) => Number(c)) : [],
        caseCodes: allNumeric ? [] : codes,
        caseText: cell(r, "case numbers"), lotCases: null,
        pkgs: cell(r, "# of pkgs"), kg: cell(r, "kg"), cbm: cell(r, "cbm"),
        shkNumber: cell(r, "client reference number"),
      };
      const declared = Number(lot.pkgs) || 0;
      if (declared && codes.length && declared !== codes.length) {
        caseCountMismatches.push({ lot: lot.lotRef || lot.unitCode || "", stated: declared, listed: codes.length });
      }
      pkgs += declared;
      kg += Number(lot.kg) || 0;
      cbm += Number(lot.cbm) || 0;

      const refJob = cell(r, "refer to job number");
      let block = refBlocks.find((b) => b.refJobNumber === refJob);
      if (!block) { block = { refJobNumber: refJob, refDate: "", shkNumber: lot.shkNumber, liftNo: lot.unitCode, lots: [] }; refBlocks.push(block); }
      block.lots.push(lot);
      if (lot.pkgs || lot.kg || lot.cbm) {
        declaredTotalsList.push({
          pkgs: lot.pkgs, kg: lot.kg, cbm: lot.cbm, refJobNumber: refJob,
          shkNumber: lot.shkNumber, context: `${lot.lotRef} ${lot.unitCode}`.trim(),
        });
      }
    }

    // Lot keys, with the same rule the job-sheet reader uses: a reference more than one lot
    // answers to identifies neither, so it is left out rather than having both lots' cases
    // merged under it.
    const owners = new Map();
    const keysOf = (lot) => [lot.lotRef, lot.unitCode].filter(Boolean).map((k) => String(k).toUpperCase());
    refBlocks.forEach((b) => b.lots.forEach((lot) => {
      new Set(keysOf(lot)).forEach((k) => owners.set(k, (owners.get(k) || 0) + 1));
    }));
    refBlocks.forEach((b) => b.lots.forEach((lot) => {
      for (const k of keysOf(lot)) {
        if ((owners.get(k) || 0) > 1) continue;
        if (lot.caseCodes.length) caseCodesByLot[k] = { codes: lot.caseCodes, text: lot.caseText };
        if (lot.caseNumbers.length) caseMarksByLot[k] = { numbers: lot.caseNumbers, text: lot.caseText, lotCases: null };
      }
    }));

    const docType = LEGACY_DOC_TYPES.includes(cell(first, "type")) ? cell(first, "type") : guessDocTypeFromName(sheet.sourceName);
    out.push({
      file, sourceName: sheet.sourceName, fromSpreadsheet: true,
      docType,
      client: (resolveClient && resolveClient(cell(first, "client"))) || "",
      projectEn: cell(first, "project"), projectZh: "",
      jobNumber: cell(first, "job number"),
      date: parseHKDate(cell(first, "date")),
      unitCode: cell(first, "lift no."),
      packageCount: pkgs ? String(pkgs) : "",
      weightKg: kg ? String(Math.round(kg * 100) / 100) : "",
      volumeCbm: cbm ? String(Math.round(cbm * 1000) / 1000) : "",
      ssDoNo: "", shkNumber: cell(first, "client reference number"),
      jobRef: cell(first, "job ref"),
      oversizeByLot: {},
      referJobNumber: cell(first, "refer to job number"), referDate: "",
      declaredTotalsList, refBlocks, caseMarksByLot, caseCodesByLot, caseCountMismatches,
      caseMarksByRef: {}, caseAutoApplied: {}, autoDetected: true,
    });
  }
  return out;
}
// Some job sheets carry their case list as a picture pasted into the sheet rather than as
// typed cells - the 2604097 CFS has all fifty-eight of its markings in a screenshot sitting
// under the C/S No. row. There is nothing in the cells to read, so the sheet looks exactly
// like one that names no cases at all, and saying "no case numbers on the sheet" sends
// someone hunting for a list that is right there in front of them.
//
// Every one of these sheets carries the letterhead as an image anchored at the top, so only
// pictures placed well down the page count. Reading the anchor needs the workbook's raw
// parts, which are kept only when the file is opened with bookFiles.
const SHEET_IMAGE_HEADER_ROWS = 20;
function sheetHasPastedContentImage(wb) {
  const files = (wb && wb.files) || {};
  for (const name of Object.keys(files)) {
    if (!/^xl\/drawings\/drawing\d*\.xml$/i.test(name)) continue;
    let xml = "";
    try {
      const f = files[name];
      xml = typeof f === "string" ? f : (f && (f.asText ? f.asText() : (f.content ? String.fromCharCode.apply(null, f.content) : "")));
    } catch (err) { continue; }
    if (!xml) continue;
    // Each anchor is taken whole, because a drawing also holds shapes and empty text boxes
    // that carry no picture at all. These sheets have several parked far down the page, and
    // counting those flagged every file as having a pasted case list. Only an anchor
    // containing a picture counts, and only one placed below the header.
    const anchors = xml.match(/<xdr:(oneCellAnchor|twoCellAnchor|absoluteAnchor)[\s\S]*?<\/xdr:\1>/g) || [];
    for (const anchor of anchors) {
      if (!/<xdr:pic[\s>]/.test(anchor) && !/r:embed=/.test(anchor)) continue;
      const from = anchor.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
      if (from && Number(from[1]) + 1 > SHEET_IMAGE_HEADER_ROWS) return true;
    }
  }
  return false;
}
// What a row has to carry before it can be processed. A date is on the list because
// everything downstream is ordered by it: the storage ledger's running balance, the free
// period, which arrival a delivery draws from, and the last-CFS column on the site summary.
// A record without one sorts to the end of the ledger and bills from nothing, and it is far
// harder to find and fix afterwards than it is to type now.
function legacyRowMissing(row) {
  const miss = [];
  if (!row.client) miss.push("client");
  if (!row.projectEn && !row.projectZh) miss.push("site");
  if (!String(row.date || "").trim()) miss.push("date");
  return miss.length ? miss : null;
}
function guessJobNumberFromName(name) {
  const m = (name || "").match(/\b(\d{6,8})\b/);
  return m ? m[1] : "";
}

// Scans a job-sheet-style Excel (Devan/CFS/Delivery layouts) for known Chinese/English
// labels and pulls whatever value sits next to each one - to the right first, then below.
// This is a best-effort pre-fill: the person still reviews and can correct every field
// before committing, same as the PDF/Excel packing-list importers already do.
const CHINESE_CLIENT_ALIASES = {
  "三菱": "Mitsubishi", "迅達": "Schindler", "蒂升": "TK Elevator", "通力": "Kone",
  "富士達": "Fujitec", "其士": "Chevalier", "希格馬": "Sigma", "日立": "Hitachi",
};
function resolveClientGuess(text) {
  if (!text) return "";
  const t = text.trim();
  for (const [zh, client] of Object.entries(CHINESE_CLIENT_ALIASES)) {
    if (t.includes(zh)) return client;
  }
  const lower = t.toLowerCase();
  const match = CLIENTS.find((c) => lower.includes(c.toLowerCase()));
  return match || "";
}
const JOBSHEET_LABEL_ALIASES = {
  account: ["客戶", "account"],
  jobNo: ["快達單號", "job no", "job no."],
  date: ["日期", "date"],
  orderedBy: ["落單人", "ordered by"],
  poNo: ["採購編號", "p.o. no", "po no"],
  jobRef: ["地盤代號", "job ref"],
  ssDoNo: ["提單資料", "ss/d.o. no", "ss/do no"],
  to: ["送", "to"],
  from: ["由", "from"],
};
const ALL_JOBSHEET_LABELS = Object.values(JOBSHEET_LABEL_ALIASES).flat();
function normCell(v) {
  return String(v == null ? "" : v).replace(/\s+/g, "").toLowerCase();
}
function normLabelCompare(v) {
  return normCell(v).replace(/[.:]+$/g, "");
}
function isKnownLabel(v) {
  const n = normLabelCompare(v);
  if (!n) return false;
  // Exact match only (after trimming trailing punctuation) - a substring check here
  // would flag real data as a label whenever a short alias like "to" happens to
  // appear inside a word (e.g. "Elevator").
  return ALL_JOBSHEET_LABELS.some((a) => n === normLabelCompare(a));
}
// Tries every cell matching one of the given labels (not just the first), and for each
// match looks to its right then below for a value - skipping any candidate that is
// itself a recognized label (adjacent field labels sit right next to each other in
// these sheets, e.g. "JOB NO." next to "DATE", so a naive "next cell" grab picks up
// the neighboring label's text instead of real data).
function findLabelValue(rows, aliases) {
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const cellNorm = normCell(rows[r][c]);
      if (!cellNorm) continue;
      // A Chinese label has to be the whole cell, not merely appear inside it. "客戶" means
      // client, and it also sits inside 柴灣水務署客戶諮詢中心 - Chai Wan Water Supplies
      // Department CUSTOMER Enquiry Centre - so a loose match read the site name as the
      // client label and returned the line underneath it. English labels keep the loose
      // match: they are written with spaces and punctuation that vary between sheets.
      if (!aliases.some((a) => {
        const alias = normCell(a);
        if (!alias) return false;
        if (/[a-z]/i.test(alias)) return cellNorm.includes(alias);
        return cellNorm === alias
          || cellNorm.replace(/[\s:\uff1a.,]+$/g, "") === alias
          || cellNorm.startsWith(`${alias} `);
      })) continue;
      for (let cc = c + 1; cc < Math.min(c + 6, rows[r].length); cc++) {
        const v = rows[r][cc];
        if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) return String(v).trim();
      }
      for (let rr = r + 1; rr < Math.min(r + 3, rows.length); rr++) {
        const v = rows[rr][c];
        if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) return String(v).trim();
      }
      // This occurrence's neighbors were empty or just more labels - keep scanning
      // the sheet in case the same label appears again (often once in Chinese, once
      // in English) with real data next to the other occurrence.
    }
  }
  return "";
}
// Collects the address block after a FROM/TO-style label - these often span several
// lines (English site name, then a separate Chinese site name below it), so this
// gathers consecutive non-empty lines and splits them by script rather than grabbing
// just the single adjacent cell.
function findAddressLines(rows, aliases, maxLines = 8) {
  const aliasesNorm = aliases.map((a) => normLabelCompare(a));
  // Prefer an exact label match over a substring match - a single Chinese character
  // alias like "送"/"由" can appear embedded inside an unrelated sentence elsewhere on
  // the sheet (e.g. "运输送到" contains "送"), which would otherwise be mistaken for
  // the real FROM/TO label cell.
  let hasExactMatch = false;
  for (const row of rows) {
    if (row.some((cell) => aliasesNorm.includes(normLabelCompare(cell)))) { hasExactMatch = true; break; }
  }
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const cellNorm = normCell(rows[r][c]);
      if (!cellNorm) continue;
      const isMatch = hasExactMatch ? aliasesNorm.includes(normLabelCompare(rows[r][c])) : aliases.some((a) => cellNorm.includes(normCell(a)));
      if (!isMatch) continue;
      let startR = -1, startC = -1;
      for (let cc = c + 1; cc < Math.min(c + 6, rows[r].length); cc++) {
        const v = rows[r][cc];
        if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) { startR = r; startC = cc; break; }
      }
      if (startR === -1) {
        for (let rr = r + 1; rr < Math.min(r + 3, rows.length); rr++) {
          const v = rows[rr][c];
          if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) { startR = rr; startC = c; break; }
        }
      }
      if (startR === -1) continue;
      const enLines = [], zhLines = [];
      for (let rr = startR; rr < Math.min(startR + maxLines, rows.length); rr++) {
        const v = rows[rr][startC];
        if (v == null || String(v).trim() === "") break;
        const text = String(v).trim();
        if (isKnownLabel(text)) break;
        if (/^(暫存|temp\.?\s*storage)/i.test(text)) break;
        if (/[\u4e00-\u9fff]/.test(text)) zhLines.push(text);
        else enLines.push(text);
      }
      if (enLines.length || zhLines.length) return { en: enLines.join(" ").trim(), zh: zhLines.join("").trim() };
    }
  }
  return { en: "", zh: "" };
}
// Converts a Date object to a plain YYYY-MM-DD string using its LOCAL calendar date,
// not toISOString() (which converts through UTC and silently shifts the date backward
// by a day for anyone in a timezone ahead of UTC, e.g. Hong Kong).
function dateToLocalISO(d) {
  const yr = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${yr}-${mo}-${da}`;
}
// Reads the workbook's own Excel "Print Area" setting for a sheet, if one is defined.
// Farspeed's rule: only what falls inside the printable A4 area counts as actually
// arriving or leaving - anything outside it was never printed or sent, so it is not part
// of this job. That applies across as well as down: the CFS_2/2605199 sheet parks a
// second, unrelated set of lots (L32-02 through L32-05, under a different SHK reference)
// out in column N, past the print area's right edge at column L, where they would
// otherwise be read as extra lots on this job and would break the lot lines they sit
// beside. Returns the last row and column covered (1-indexed, as Excel counts them).
function colLettersToIndex(letters) {
  let n = 0;
  for (const ch of String(letters || "").toUpperCase()) {
    if (ch < "A" || ch > "Z") return 0;
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}
function getPrintArea(wb, sheetIndex) {
  try {
    const names = wb.Workbook && wb.Workbook.Names;
    if (!names) return null;
    for (const n of names) {
      if (!n || !n.Name || !/print_area/i.test(n.Name)) continue;
      if (n.Sheet !== undefined && n.Sheet !== null && n.Sheet !== sheetIndex) continue;
      const ref = String(n.Ref || "");
      const m = ref.match(/!\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?/);
      if (!m) continue;
      const lastRow = Math.max(Number(m[2]), Number(m[4] || m[2]));
      const lastCol = Math.max(colLettersToIndex(m[1]), colLettersToIndex(m[3] || m[1]));
      return { lastRow, lastCol: lastCol || 0 };
    }
  } catch (e) { /* fall back to using the whole sheet */ }
  return null;
}
// Job sheets write the referral date free-hand and day-first ("Ref Job no. 2605126 on
// 16/5/2026"), which Date() reads month-first - it rejects 16/5 outright and would turn
// 5/6 into the wrong day without complaining. Separators are inconsistent too ("29- /5/
// 2026" is a real line from this file), so they are normalised before the parts are read
// off directly.
// A day written as a span rather than a single date - "10-13/5/2026" for work running over
// several days, or "7-/5/2026" where the writer meant "the 7th onwards" and never came back
// to finish it. The job happened on the first of those days: that is when the goods moved,
// when storage starts, and where the movement belongs in the ledger. So the earlier number
// is taken and the rest of the span dropped.
function firstDayOfSpan(raw) {
  return String(raw == null ? "" : raw)
    .replace(/\s+/g, "")
    // The span must be followed by a slash or a dot, which is how these sheets write them:
    // "10-13/5/2026", "1-3.6.2026". Allowing a hyphen there swallowed ordinary dates -
    // "3-2-26" is the 3rd of February, not a span - and left them to be misread.
    .replace(/^(\d{1,2})\s*-\s*(\d{0,2})(?=[\/\.])/, "$1");
}
// Every date this software reads off a file is a Hong Kong date. Where the value is a real
// date - a Date object out of a spreadsheet cell - it is taken as it stands, because there
// is nothing to interpret. Where it is text, it is read day-month-year: "07/01/2026" is the
// 7th of January. Letting the browser decide instead means American order, which is wrong
// here and wrong silently, since it only differs on the twelve days a month where both
// readings are possible.
function parseHKDate(value) {
  if (value instanceof Date && !isNaN(value)) return dateToLocalISO(value);
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";
  // Already year-first and unambiguous.
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const dayFirst = parseSheetDayFirstDate(text);
  if (dayFirst) return dayFirst;
  // Something else entirely - a month name, say - where there is no day/month order to get
  // wrong.
  const d = new Date(text);
  return isNaN(d) ? "" : dateToLocalISO(d);
}
function parseSheetDayFirstDate(raw) {
  const s = firstDayOfSpan(raw).replace(/[\/\.\-]+/g, "/");
  if (!s) return "";
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return "";
  const yr = m[3].length === 2 ? `20${m[3]}` : m[3];
  const day = Number(m[1]), mon = Number(m[2]);
  if (day < 1 || day > 31 || mon < 1 || mon > 12) return "";
  return `${yr}-${String(mon).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}
// A case marking, as Irene writes them - and the separator carries meaning:
//   "#3-7/34"   a hyphen is an inclusive range: cases 3,4,5,6,7 of a 34-case lot
//   "#3,7/19"   a comma is a discrete list:     cases 3 and 7 only, of a 19-case lot
//   "1-16/16"   the whole lot, written in the C/S NO. column rather than on its own line
// The "/34" tail is the lot's case count, not another case number, so it is taken off
// first. Cells typed with a leading apostrophe to stop Excel reformatting them keep that
// apostrophe once parsed ("'60766021/L52#3,7/19"), so it is stripped as well.
function parseCaseMark(text) {
  const original = String(text == null ? "" : text).trim().replace(/^['\u2018\u2019]/, "").trim();
  let s = original.replace(/^#/, "").trim();
  if (!s) return null;
  let lotCases = null;
  const slash = s.lastIndexOf("/");
  if (slash > -1) {
    const tail = s.slice(slash + 1).trim();
    if (/^\d+$/.test(tail)) { lotCases = Number(tail); s = s.slice(0, slash); }
  }
  s = s.replace(/#/g, "").trim();
  if (!s || !/^[\d\s,\-]+$/.test(s)) {
    // Some sheets repeat the lot size after every case instead of once at the end -
    // "1/2,2/2" rather than "1,2/2" - which the reading above cannot take. parseCaseSpec
    // handles that shape, so it is the fallback rather than giving up on the marking.
    const spec = parseCaseSpec(original);
    return spec.numbers.length ? { numbers: spec.numbers, lotCases: spec.lotCases, text: original } : null;
  }
  const numbers = [...parseRangeInput(s)].sort((a, b) => a - b);
  if (!numbers.length) return null;
  return { numbers, lotCases, text: original };
}
function mergeLotMark(lot, mark) {
  if (!mark) return;
  lot.caseNumbers = [...new Set([...(lot.caseNumbers || []), ...mark.numbers])].sort((a, b) => a - b);
  if (!lot.caseText) lot.caseText = `#${mark.text.replace(/^#/, "")}`;
  if (lot.lotCases == null) lot.lotCases = mark.lotCases;
}
// One delivery job sheet can close out more than one arrival. Schindler's 2606033 sheet
// carries "Ref Job no. 2605126" with a single lot, then "Ref Job no. 2605199" with two
// more below it, and only then one "共:" line covering all eight packages between them.
// Read as a single flat delivery the second arrival disappears entirely, and the figures
// each block declares for itself (2,208kg/11.44cbm against the first, 2,228kg/6.80cbm
// against the second) are lost behind the combined total. So the page is split here into
// one block per "Ref Job no." line, each holding its own lots, case numbers and totals.
// One token of a written case list: "8B-1", "01C3101-4-1", "(10-1/10B-1)". Brackets and
// slashes are part of a marking, not punctuation around it.
//
// A marking is also written with a number sign in front of it - "#01C01", "#07D5107" - and
// Excel leaves a stray apostrophe behind wherever someone typed one to force a cell to
// text: "07C21,'08C22,08C23". Neither is part of the marking, but both have to be allowed
// here or the line carrying them stops looking like a case list at all. The 2608058 CFS
// lost ninety-eight of its hundred and twenty cases to a single apostrophe halfway down.
const CASE_CODE_LEAD = "[#'\\u2018\\u2019]?";
const CASE_CODE_TOKEN = `${CASE_CODE_LEAD}[0-9A-Za-z(][0-9A-Za-z\\-\\/\\.()]*`;
// Taken off again before the marking is stored, so what is kept is what is painted on the
// case and what the packing list holds.
function cleanCaseCode(c) {
  return String(c == null ? "" : c).trim().replace(/^[#'\u2018\u2019]+/, "").trim();
}
// A full stop typed where a comma was meant: "09B11.09B81,09D41" on the 2608191 delivery,
// which read as one case called 09B11.09B81 and left that lot a case short of the twenty-
// seven it declares - and the case it did produce matches nothing at the depot, because no
// such marking exists.
//
// The stop is only taken as a separator when both sides look like markings in their own
// right: at least three characters with a letter among them. A decimal such as "1.5" splits
// into "1" and "5", neither of which carries a letter, so it is left alone - and no marking
// on any sheet here contains a full stop, so nothing legitimate is being broken up.
function splitStrayFullStops(token) {
  const parts = String(token || "").split(".");
  if (parts.length < 2) return [token];
  const looksLikeMarking = (t) => t.length >= 3 && /[A-Za-z]/.test(t) && /\d/.test(t);
  return parts.every(looksLikeMarking) ? parts : [token];
}
// A hand-typed list separates its last pair with an ampersand rather than a comma - "01B81,
// 01D41, 01Z11, 91B11, 02B71, 02D41 & 92B11" - and a Chinese keyboard reaches for the
// ideographic comma. Both are commas as far as the list is concerned.
const CASE_CODE_SEP = "\\s*[,&\\u3001]\\s*";
const CASE_CODE_LIST_RE = new RegExp(`^${CASE_CODE_TOKEN}(${CASE_CODE_SEP}${CASE_CODE_TOKEN})*\\s*[,&\\u3001]?$`);
const CASE_CODE_SPLIT_RE = /[,&\u3001]/;
// Two markings are the same case when they read the same once spacing and the brackets a
// sheet puts round a combined package are taken off: a job sheet writes "(10-1/10B-1)" for
// the case the packing list holds as "10-1/10B-1".
// A job sheet and its packing list do not always write a case to the same length. The
// 2607140 delivery asks for "03B11" - lift and component - where the packing list holds
// "03B1103", the same case with its own number on the end. Compared whole, nothing matches
// and every case has to be ticked by hand.
//
// So an exact match is tried first, and only then a case whose marking begins with what the
// sheet asked for. That fallback has to be unique to be trusted: where a lot holds
// 03B4203-2-1 and 03B4203-2-2, a sheet asking for "03B42" could mean either, and guessing
// would put the wrong case out of the door. Ambiguous ones are left unmatched and reported.
function resolveCaseCode(code, packages) {
  const exact = (packages || []).find((p) => sameCaseCode(p.code, code));
  if (exact) return exact;
  const norm = (c) => String(c || "").toUpperCase().replace(/[\s()#'\u2018\u2019]/g, "");
  const wanted = norm(code);
  if (wanted.length < 4) return null;
  const starts = (packages || []).filter((p) => norm(p.code).startsWith(wanted));
  return starts.length === 1 ? starts[0] : null;
}
function sameCaseCode(a, b) {
  const norm = (c) => String(c || "").toUpperCase().replace(/[\s()#'\u2018\u2019]/g, "");
  return norm(a) === norm(b);
}
// A case list typed as a grid rather than a line: each cell one marking, read left to right
// then down. Every cell has to look like a marking on its own, or an ordinary row of a job
// sheet would be swallowed whole.
const SINGLE_CASE_CODE_RE = new RegExp(`^${CASE_CODE_TOKEN}$`);
function gridCaseCells(cells) {
  const vals = (cells || []).filter(Boolean);
  if (vals.length < 2) return null;
  for (const v of vals) {
    if (!SINGLE_CASE_CODE_RE.test(v)) return null;
    if (!/[A-Za-z]/.test(v) || !/\d/.test(v)) return null;
  }
  return vals.flatMap((v) => splitStrayFullStops(cleanCaseCode(v))).map(cleanCaseCode).filter(Boolean);
}
// A case list written out in full keeps the form it was written in. Turning "02, 03, 061,
// 081" into the numbers 2, 3, 61, 81 throws away the padding that is painted on the case and
// printed on the packing list, and a marking that no longer matches what the depot holds is
// where these mistakes come from. Ranges still expand to numbers - "1-16/16" has to - but a
// list naming each case, with a leading zero anywhere in it, is kept verbatim.
function explicitZeroPaddedCodes(text) {
  const tokens = String(text || "").replace(/^#/, "").split(/[,\s]+/).map((t) => t.trim()).filter(Boolean);
  if (!tokens.length) return null;
  if (!tokens.every((t) => /^\d+$/.test(t))) return null;
  // A lone number is a count far more often than it is a case, so one on its own only
  // counts when it is padded - "01", "07" on their own lines under an escalator's C/S row -
  // or when the lot it belongs to is already collecting written markings.
  if (!tokens.some((t) => /^0\d/.test(t))) return null;
  return tokens;
}
function parseJobSheetBlocks(rows) {
  const blocks = [];
  let cur = null;
  // A SHK reference or LIFT NO. line sits above the lots it describes and can be restated
  // partway down the page - the 2605199 CFS sheet files L52/L53 under SHK0395/26 and then
  // L32-02..05 under SHK0107/26 - so the heading in force is tracked as the page is read
  // rather than being pinned to whichever block happened to be open.
  let ctx = { shkNumber: "", liftNo: "" };
  // Set when a C/S NO. row stated its figures but named no cases, so the line under it may
  // be carrying that lot's case list without a leading "#".
  let awaitingMark = false;
  // Counts C/S NO. rows carrying a weight that could not be placed on a lot because that
  // lot had already stated one. A sheet whose C/S rows are per-case rather than per-lot
  // produces these in quantity, and that is the signal not to trust per-lot figures.
  let orphanFigureRows = 0;
  // Set once the page has added itself up. What the sheet refers to after that point is
  // about the job as a whole rather than another arrival being drawn from.
  let sawGrandTotal = false;
  // The delivery's own Delivery Memo number, written at the foot as "Refer to 13-DM-26-0228".
  let ownDmRef = "";
  // Set by a LIFT NO. line: the row under it may name the lot those lifts belong to.
  let awaitingLotName = false;
  // Set by a C/S NO. row: the rows under it may continue the same group's figures without
  // repeating the label.
  let awaitingFigures = false;
  // A block counts as closed once figures have been stated for it - either over the block
  // as a whole, or on the last lot it named. Lot-level figures have to count, or a heading
  // restated further down the page would be read as belonging to the lots above it.
  const closed = (b) => {
    if (!b) return false;
    if (b.pkgs || b.kg || b.cbm) return true;
    const last = b.lots[b.lots.length - 1];
    return !!(last && (last.pkgs || last.kg || last.cbm));
  };
  const startBlock = (refJobNumber, refDateRaw) => {
    cur = {
      refJobNumber: refJobNumber || "", refDateRaw: refDateRaw || "",
      shkNumber: ctx.shkNumber, liftNo: ctx.liftNo, lots: [], pkgs: "", kg: "", cbm: "",
      figuresAfterLots: false,
    };
    blocks.push(cur);
    return cur;
  };
  const allRows = rows || [];
  // The first cell of the next row that has anything in it - used to recognise a lot
  // heading by what follows it rather than by what precedes it.
  const nextMeaningful = (from) => {
    for (let k = from + 1; k < allRows.length; k++) {
      const cs = (allRows[k] || []).map((c) => String(c == null ? "" : c).trim());
      if (cs.some(Boolean)) return cs.find(Boolean) || "";
    }
    return "";
  };
  for (let ri = 0; ri < allRows.length; ri++) {
    const raw = allRows[ri];
    const cells = (raw || []).map((c) => String(c == null ? "" : c).trim());
    const line = cells.filter(Boolean).join(" ").trim();
    if (!line) {
      // A blank row ends a case list that has already started. The list itself may run over
      // several rows with no gap between them, but once it stops, it stops.
      const openLot = cur && cur.lots[cur.lots.length - 1];
      if (openLot && (openLot.caseCodes || []).length) awaitingMark = false;
      continue;
    }

    const refM = line.match(/ref(?:er)?\.?\s*(?:to\s+)?job\s*no\.?\s*([A-Za-z0-9\-]+)\s*(?:on\s*([\d\/\.\- ]+\d))?/i)
      // "Refer to 13-DM-26-0228", with the words "job no." left out. Only a reference that
      // is unmistakably a lot is read this way, or every sentence beginning "refer to"
      // would open a block of its own.
      || line.match(/refer(?:red)?\s*to\s*(\d{1,3}\s*-\s*DM\s*-\s*\d{2}\s*-\s*\d{0,4})\s*(?:on\s*([\d\/\.\- ]+\d))?/i);
    if (refM) {
      ctx = { shkNumber: "", liftNo: "" };
      const ref = refM[1].trim();
      const refDate = (refM[2] || "").trim();
      // A delivery refers back to the arrival job it is drawing from, and names the lot it
      // is taking out of it. The two are different things, and the sheets say "Refer to job
      // no." for both: "Refer to job no. 2602081" is the arrival, while "Refer to job no.
      // 13-DM-25-0616" is a lot - Mitsubishi's Delivery Memo number, where Schindler would
      // write an SHK reference and another maker something else again. Filed as a job
      // number it was hunted for among arrival job numbers, which it will never be, and the
      // block came through with no lot at all. Anything that is not a plain job number is
      // taken as the lot's reference instead.
      if (/^\d{6,8}$/.test(ref)) {
        startBlock(ref, refDate);
      } else if (sawGrandTotal) {
        // Below the total, a Delivery Memo number is the one this delivery was issued
        // under - "Refer to 13-DM-26-0228", or "Refer to 13-DM-26-0312, 0310, 0301" where
        // several cover it. It is the sheet's own reference, not another lot going out, and
        // reading it as a lot put a phantom row carrying no packages, weight or cases on
        // the end of every delivery sheet in the backlog.
        ownDmRef = ownDmRef || line.replace(/^.*?refer(?:red)?\s*(?:to\s*)?(?:job\s*no\.?)?\s*/i, "").trim();
        continue;
      } else {
        startBlock("", refDate);
        cur.lots.push({
          lotRef: ref.replace(/\s+/g, ""), altRef: "", unitCode: "",
          caseNumbers: [], caseCodes: [], caseText: "", lotCases: null,
          pkgs: "", kg: "", cbm: "", shkNumber: "",
        });
        awaitingMark = true;
      }
      continue;
    }

    // A 共:/Total: line closes the block it belongs to; the grand total at the foot of the
    // page belongs to no single arrival and must not overwrite the last block's figures.
    if (/(?:共|total)\s*[:\uff1a]/i.test(line)) { cur = null; ctx = { shkNumber: "", liftNo: "" }; sawGrandTotal = true; continue; }

    const shkM = line.match(/\bSHK\s*[-#]?\s*(\d{3,6})\s*\/\s*(\d{2})\b/i);
    // "LIFT NO. L-W07, L-W08" names two lifts and "#01 & #02" two more; taking only the
    // first lost the rest, so the whole tail of the line is kept.
    const liftM = line.match(/lift\s*no\.?\s*[:\-]?\s*(.+)$/i);
    if (shkM || liftM) {
      // A heading arriving after the open block has already stated its figures belongs to
      // what comes next, not to what just finished.
      if (closed(cur)) cur = null;
      if (shkM) { ctx.shkNumber = `SHK${shkM[1]}/${shkM[2]}`; if (cur) cur.shkNumber = ctx.shkNumber; }
      if (liftM) {
        ctx.liftNo = liftM[1].replace(/#/g, "").replace(/\s+/g, " ").trim();
        awaitingLotName = true;
        if (cur) cur.liftNo = ctx.liftNo;
      }
      // A heading line is nothing else, and falling through would clear the flag it just
      // set before the row below could be read as the lot name.
      continue;
    }

    // Mitsubishi names a lot by its Delivery Memo number rather than an order/lift pair:
    // "DM No. | 13-DM-26-0500 | S/M: | 1325003000" on one row, with the lift given above as
    // "LIFT NO. #01-02". That DM number is what Irene files the cases under, so it becomes
    // the lot reference and the shipping mark is kept alongside it.
    if (/^DM\s*No\.?/i.test(cells.find(Boolean) || "")) {
      const rest = cells.filter(Boolean).slice(1);
      const dmNo = (rest.find((c) => /\d{1,3}\s*-\s*DM\s*-\s*\d{2}\s*-\s*\d{3,4}/i.test(c)) || "").trim();
      const smIdx = rest.findIndex((c) => /^S\/M/i.test(c));
      const shippingMark = smIdx >= 0 ? (rest[smIdx + 1] || "").trim() : "";
      // A DM row is sometimes written before its number is issued - "DM No. 13-DM-26-" with
      // the tail blank. The shipping mark beside it still identifies the lot, so the row is
      // not thrown away for want of the number.
      if (dmNo || shippingMark) {
        if (closed(cur) && cur.figuresAfterLots) startBlock(cur.refJobNumber, cur.refDateRaw);
        if (!cur) startBlock("", "");
        cur.lots.push({
          lotRef: (dmNo || shippingMark).replace(/\s+/g, ""),
          altRef: dmNo ? shippingMark : "", unitCode: ctx.liftNo || "",
          // Set when the DM number was missing and the mark had to stand in for it. A mark
          // covers a whole contract rather than one lot, so this is checked below.
          markOnly: !dmNo,
          caseNumbers: [], caseCodes: [], caseText: "", lotCases: null,
          pkgs: "", kg: "", cbm: "", shkNumber: ctx.shkNumber,
        });
        // A delivery sheet states the lot's figures on the DM row itself - "DM No. |
        // 13-DM-26-0073 | 10 | PKGS | ELEVATOR MATERIALS | 1994 | KGS | 7.86 | CBM" -
        // rather than on a C/S NO. row of its own. Pushing the lot and moving on left every
        // one of those lots with no packages, no weight and no volume at all.
        const dmLot = cur.lots[cur.lots.length - 1];
        const dmPk = line.match(/(\d+)\s*PKGS?/i);
        const dmKg = line.match(/([\d,]+(?:\.\d+)?)\s*KGS?\b/i);
        const dmCb = line.match(/([\d,]+(?:\.\d+)?)\s*CBM\b/i);
        if (dmPk) dmLot.pkgs = dmPk[1];
        if (dmKg) dmLot.kg = dmKg[1].replace(/,/g, "");
        if (dmCb) dmLot.cbm = dmCb[1].replace(/,/g, "");
        if (dmLot.kg || dmLot.cbm) cur.figuresAfterLots = true;
        // The list under it still has to be read, so stay open unless this row named cases.
        awaitingMark = !!(dmPk || dmKg || dmCb);
        continue;
      }
    }
    // "60759188/L32-01", with its cases written inline ("'60766021/L52#3,7/19"), and/or its
    // own weight and volume trailing on the same line - "60778397/L34-02  3018.58 KGS  6.36
    // CBM" - which is how a sheet states figures per lot without a C/S NO. row for each.
    const lotM = line.match(/^['\u2018\u2019]?([A-Za-z]{0,4}\d{6,12})(?:\s*\/\s*([A-Za-z]{0,4}\d{6,12}))?\s*\/\s*([A-Za-z][A-Za-z0-9\-\.]*)\s*(?:#\s*([\d][\d\s,\-\/]*?))?\s*(?:([\d,]+(?:\.\d+)?)\s*KGS?\b)?\s*(?:([\d,]+(?:\.\d+)?)\s*CBM\b)?\s*$/i);
    if (lotM) {
      // A CFS sheet names a lot and then closes it with its own C/S NO. figures before
      // naming the next, so a lot arriving after a block was closed that way starts a new
      // one. A delivery sheet does the opposite - it states the figures first and then
      // lists the lots they cover - so the test is whether the figures were recorded with
      // lots already open, not merely whether figures exist.
      if (closed(cur) && cur.figuresAfterLots) startBlock(cur.refJobNumber, cur.refDateRaw);
      if (!cur) startBlock("", "");
      // A lot may be named by a container/UID reference, by an order number, or by both:
      // "60737177/P3", "HKG0011764678/P3P4", "HKG0011764688/ 60748116/P3". Where both are
      // present the order number is the identifying one and the UID is kept alongside it,
      // since either may be what the packing list was filed under.
      const lot = {
        lotRef: lotM[2] || lotM[1], altRef: lotM[2] ? lotM[1] : "", unitCode: lotM[3],
        caseNumbers: [], caseText: "", lotCases: null,
        pkgs: "", kg: lotM[5] ? lotM[5].replace(/,/g, "") : "", cbm: lotM[6] ? lotM[6].replace(/,/g, "") : "",
        shkNumber: ctx.shkNumber,
      };
      if (lotM[4]) mergeLotMark(lot, parseCaseMark(lotM[4]));
      cur.lots.push(lot);
      if (lot.kg || lot.cbm) cur.figuresAfterLots = true;
      awaitingMark = false;
      awaitingLotName = false;
      continue;
    }
    // Chevalier heads each group with its own line naming the lifts and the order it
    // covers - "CED1832B", or "L-C01 to L-C05/CED-1833/B". There is no fixed shape to it,
    // so it is recognised by position: the line immediately above a C/S NO. row, once
    // everything with a shape of its own has had its turn - a referral, a DM row, a
    // numbered lot line and a figures row are all matched above this. One LIFT NO. line
    // can cover several such groups, so keying off the lifts alone found only the first.
    if (/^C\/S\s*NO\.?/i.test(nextMeaningful(ri))
        && !/^C\/S/i.test(line) && !/^DM\s*No/i.test(line) && !/S\/M/i.test(line)
        && !/(PKGS?|KGS?|CBM)\b/i.test(line)
        && line.length <= 48 && /[A-Za-z]/.test(line) && /\d/.test(line)) {
      if (closed(cur) && cur.figuresAfterLots) startBlock(cur.refJobNumber, cur.refDateRaw);
      if (!cur) startBlock("", "");
      cur.lots.push({
        lotRef: line, altRef: "", unitCode: ctx.liftNo || "",
        caseNumbers: [], caseCodes: [], caseText: "", lotCases: null,
        pkgs: "", kg: "", cbm: "", shkNumber: ctx.shkNumber,
      });
      awaitingLotName = false;
      awaitingMark = false;
      continue;
    }
    awaitingLotName = false;
    if (!cur) continue; // ordinary header rows above the first lot or referral

    // A marking on a line of its own belongs to the lot named directly above it. This is
    // the numeric form - "#1-16/16", "#1-8,14,16-21" - and it is claimed only when it
    // actually parses as one. These sheets also number their markings with a hash,
    // "#01C01, #04C10" and "#07D5107" on a line by itself, and claiming those here as well
    // swallowed them: nothing numeric could be made of them, so the line was consumed and
    // thrown away, and thirty-six cases read as none.
    if (/^#/.test(line)) {
      // "#1-16/16" and "#1-8,14" are markings. "#03" alone, sitting above its own C/S NO.
      // row, is a lift heading - the 2604119 Devan writes "#03" and "#04" that way - and
      // reading it as case number 3 gave that lot two cases against the twenty it declares
      // and left the real lists below it unread. A single number with no range, no list and
      // no C/S row waiting on it is a heading.
      const looksLikeList = /[,\-\/]/.test(line);
      const mark = looksLikeList ? parseCaseMark(line) : null;
      if (mark && (mark.numbers || []).length) {
        if (cur.lots.length) mergeLotMark(cur.lots[cur.lots.length - 1], mark);
        awaitingMark = false;
        continue;
      }
    }
    // Some sheets write that list without the "#", on the line under a C/S NO. row that
    // stated no cases itself - "1-8,14,16-21,26,28-32" beneath L34-S1's row. A bare list of
    // numbers is only read this way in that exact position, so an ordinary figure elsewhere
    // on the page is never mistaken for a case marking.
    //
    // An escalator sheet puts the cases in the first column and what is in them in another,
    // a row at a time: "02, 03, 04, 05, 06, 061, 08, 081 | ESC. PARTS", then "01 | UPPER
    // TRUSS ASS'Y", then "08 (與A2疊起) | CENTRE TRUSS ASS'Y". Joined for matching, every one
    // of those rows carries its description along and stops looking like a case list, so
    // the 2602111 Devan read none of its eighteen. The first cell alone is what is offered
    // here, with any note in brackets after the marking taken off, and the list stays open
    // for the rows below it - a case named twice, once in the run and again on its own line
    // with a note, is still one case.
    const firstCell = cells.find(Boolean) || "";
    const markCell = firstCell.replace(/[\(\uff08][^)\uff09]*[\)\uff09]\s*$/, "").trim();
    // The first cell only stands in for the whole row when the row is not itself a grid of
    // markings. On a grid row every cell is a case, and reading just the first turned the
    // 2601009 CFS's forty-four into eleven - one per row instead of four.
    const isGridRow = !!gridCaseCells(cells);
    const listCandidates = !isGridRow && markCell && markCell !== line ? [line, markCell] : [line];
    for (const candidate of listCandidates) {
      if (!awaitingMark || !/^[\d][\d\s,\-]*$/.test(candidate)) continue;
      const lot = cur.lots[cur.lots.length - 1];
      const padded = explicitZeroPaddedCodes(candidate);
      if (lot && padded) {
        const seen = new Set((lot.caseCodes || []).map((c) => String(c).toUpperCase()));
        const fresh = padded.filter((c) => !seen.has(c.toUpperCase()));
        lot.caseCodes = [...(lot.caseCodes || []), ...fresh];
        lot.caseText = lot.caseCodes.join(", ");
        awaitingMark = true;
        break;
      }
      if (lot) mergeLotMark(lot, parseCaseMark(candidate));
      // Stay open: these lists run a row at a time rather than all on one line.
      awaitingMark = true;
      break;
    }
    // A row can carry the list in its first column and that lot's figures in the rest -
    // "02, 03, 04, 05, 06, 061 | 8 | PKGS | 6057 | KGS | 38.939 | CBM". Taking the list and
    // moving on threw the figures away, which is why that lot showed no packages, weight or
    // volume at all. The row is left to carry on to the figures below.
    const rowAlsoStatesFigures = /\d+\s*PKGS?/i.test(line) && /(KGS?|CBM)\b/i.test(line);
    if (awaitingMark && !rowAlsoStatesFigures && listCandidates.some((c) => /^[\d][\d\s,\-]*$/.test(c))) continue;
    // Mitsubishi writes its cases as full markings rather than numbers - "01A2101,
    // 02A2102, 01C3101-4-1" - wrapped across as many lines as it takes, each continuing
    // line following on from the trailing comma of the one above. They are codes, not
    // numbers, so they are kept as written and matched whole.
    //
    // A marking can also name two cases that travel as one package, bracketed and split by
    // a slash - "(10-1/10B-1)", "(9A-1/9B-1)". Neither character was allowed here, so a
    // line holding one failed to be a case list at all and the whole block came through
    // with no cases: the 2607208 delivery lists eleven that way and read as none.
    const codeLine = listCandidates.find((c) => CASE_CODE_LIST_RE.test(c) && /[A-Za-z]/.test(c) && /\d/.test(c));
    if (awaitingMark && codeLine && !rowAlsoStatesFigures) {
      const codes = codeLine.split(CASE_CODE_SPLIT_RE)
        .flatMap((c) => splitStrayFullStops(cleanCaseCode(c)))
        .map(cleanCaseCode).filter(Boolean);
      const lot = cur.lots[cur.lots.length - 1];
      if (lot) {
        lot.caseCodes = [...(lot.caseCodes || []), ...codes];
        lot.caseText = lot.caseCodes.join(", ");
      }
      // Stay open for the row below. The list runs on for as many rows as it takes and the
      // trailing comma is not reliable - the 2606062 Devan breaks its twenty cases into two
      // rows of ten with no comma at the end of the first, so keying on that read only the
      // first lift's ten. A blank row or any other kind of line ends it instead.
      awaitingMark = true;
      continue;
    }
    // The same list is sometimes typed into a grid instead of a line: the 2601009 CFS puts
    // its forty-four markings one to a cell across four columns and eleven rows, with no
    // commas anywhere. Joined up for matching that reads "23E2123 23D4123 24E2124 24D4124",
    // which is not a comma-separated list, so the block came through with no cases at all
    // against the forty-four it declares.
    //
    // A grid row counts only when every cell on it is a marking in its own right - letters
    // and digits both, nothing that is a label or a bare number - and there are at least
    // two of them. A single cell is indistinguishable from a lot name or a stray heading.
    if (awaitingMark) {
      const gridCodes = gridCaseCells(cells);
      if (gridCodes) {
        const lot = cur && cur.lots[cur.lots.length - 1];
        if (lot) {
          lot.caseCodes = [...(lot.caseCodes || []), ...gridCodes];
          lot.caseText = lot.caseCodes.join(", ");
        }
        awaitingMark = true;
        continue;
      }
    }
    awaitingMark = false;

    // The C/S NO. row states the block's own package/weight/volume figures, and on some
    // sheets the case marking too, sitting in its own column: "C/S NO. | 1-16/16 | 16 | PKGS".
    // A group's figures often run over several rows - Chevalier writes one line per rail
    // type and a third for the fastening hardware - with only the first carrying the
    // "C/S NO." label. Rows under it that state PKGS and a weight continue the same group,
    // and the figures are summed rather than the first winning: CED1832B's 6 + 6 + 1
    // packages and 2,058 + 4,554 + 87 kg are the lot, not its first line.
    const isCsRow = /^C\/S\s*NO\.?/i.test(cells.find(Boolean) || "");
    const isFigureRow = /\d+\s*PKGS?/i.test(line) && /(KGS?|CBM)\b/i.test(line);
    if (!isCsRow && !isFigureRow) awaitingFigures = false;
    if (isCsRow || (awaitingFigures && isFigureRow)) {
      // The marking sits in a cell of its own before the package count. Scanning stops at
      // the PKGS column so that count can never be read as a case number of its own.
      const pkgsIdx = cells.findIndex((c) => /PKGS?/i.test(c));
      const scanEnd = pkgsIdx > 0 ? pkgsIdx : cells.length;
      // Whether this row carried its own marking, which decides below whether the rows
      // under it still need reading.
      let markedHere = false;
      for (let ci = 1; ci < scanEnd; ci++) {
        const cell = cells[ci];
        if (!cell) continue;
        // Either "1-16/16", which states the lot size, or a plain range like "4-9". A
        // plain one has to name at least two cases: a lone number in this position is a
        // count, not a marking, which is why the lot size used to be required outright -
        // at the cost of the 2607208 sheet's "4-9" being read as no cases at all.
        const withLotSize = /^#?\s*\d[\d\s,\-]*\/\s*\d+[\d\s,\-\/]*$/.test(cell);
        const plainList = /^#?\s*\d[\d\s,\-]*$/.test(cell) && /[,\-]/.test(cell);
        // The row can equally carry the markings themselves - "C/S No. |
        // 01A11,01B11,02B11,..." - which is how these delivery sheets write them, with the
        // rest of the list continuing on the rows beneath. Only the continuation rows were
        // being read, so a lot whose list fitted on one line came through with nothing and a
        // longer one came through missing its first line.
        // At least two markings, so a lone reference in this cell is not mistaken for a
        // case: the 2604042 delivery writes the DM number it is drawing against here -
        // "C/S No. | 13-DM-25-0625 | 10 | PKGS" - and reading that as a case gave the lot
        // eleven against the ten it declares.
        const codeList = !withLotSize && !plainList
          && CASE_CODE_SPLIT_RE.test(cell)
          && !/\d{1,3}\s*-\s*DM\s*-\s*\d{2}/i.test(cell)
          && CASE_CODE_LIST_RE.test(cell) && /[A-Za-z]/.test(cell) && /\d/.test(cell);
        if (!withLotSize && !plainList && !codeList) continue;
        if (codeList) {
          const lotForCodes = cur.lots[cur.lots.length - 1];
          if (lotForCodes) {
            const found = cell.split(CASE_CODE_SPLIT_RE)
              .flatMap((c) => splitStrayFullStops(cleanCaseCode(c)))
              .map(cleanCaseCode).filter(Boolean);
            const seenCodes = new Set((lotForCodes.caseCodes || []).map((c) => String(c).toUpperCase()));
            lotForCodes.caseCodes = [...(lotForCodes.caseCodes || []), ...found.filter((c) => !seenCodes.has(c.toUpperCase()))];
            lotForCodes.caseText = lotForCodes.caseCodes.join(", ");
          }
          // Deliberately not marked as finished: the list carries on below.
          break;
        }
        // Same rule as a list written on the rows below: an explicit padded list keeps the
        // form it was written in rather than being flattened to numbers.
        const lotHere = cur.lots[cur.lots.length - 1];
        const paddedHere = withLotSize ? null : explicitZeroPaddedCodes(cell);
        if (lotHere && paddedHere) {
          const seen = new Set((lotHere.caseCodes || []).map((c) => String(c).toUpperCase()));
          lotHere.caseCodes = [...(lotHere.caseCodes || []), ...paddedHere.filter((c) => !seen.has(c.toUpperCase()))];
          lotHere.caseText = lotHere.caseCodes.join(", ");
        } else if (lotHere) {
          mergeLotMark(lotHere, parseCaseMark(cell));
        }
        markedHere = true;
        break;
      }
      const pk = line.match(/(\d+)\s*PKGS?/i);
      const kg = line.match(/([\d,]+(?:\.\d+)?)\s*KGS?/i);
      const cb = line.match(/([\d,]+(?:\.\d+)?)\s*CBM/i);
      const kgVal = kg ? kg[1].replace(/,/g, "") : "";
      const cbVal = cb ? cb[1].replace(/,/g, "") : "";
      const lot = cur.lots[cur.lots.length - 1];
      // A C/S NO. row under a lot states that lot's own figures. Where the row instead
      // comes before any lot is named - a delivery sheet gives the figures first and then
      // lists the lots they cover - it belongs to the block as a whole.
      const add = (was, more, dp) => {
        if (!more) return was;
        const sum = (Number(was) || 0) + Number(more);
        return String(Math.round(sum * Math.pow(10, dp)) / Math.pow(10, dp));
      };
      if (lot) {
        lot.pkgs = add(lot.pkgs, pk ? pk[1] : "", 0);
        lot.kg = add(lot.kg, kgVal, 2);
        lot.cbm = add(lot.cbm, cbVal, 3);
        if (lot.kg || lot.cbm) cur.figuresAfterLots = true;
      } else if (kgVal || cbVal) {
        cur.pkgs = add(cur.pkgs, pk ? pk[1] : "", 0);
        cur.kg = add(cur.kg, kgVal, 2);
        cur.cbm = add(cur.cbm, cbVal, 3);
      } else if (pk) cur.pkgs = add(cur.pkgs, pk[1], 0);
      awaitingFigures = true;
      const hadFigures = closed(cur);
      if (!hadFigures && closed(cur) && cur.lots.length > 0) cur.figuresAfterLots = true;
      const lastLot = cur.lots[cur.lots.length - 1];
      // Armed whenever this C/S NO. row named no cases of its own. It used to check whether
      // the lot already held any, which stopped a second C/S NO. row for the same lot from
      // being read at all: the 2604119 Devan states "#03 | C/S NO. 10 PKGS" with its ten
      // markings, then "#04 | C/S NO. 10 PKGS" with ten more, and only the first ten came
      // through against the twenty the sheet totals. Anything that is not a case list
      // closes it again on the very next row, so re-arming costs nothing.
      awaitingMark = !!lastLot && !markedHere;
      continue;
    }
  }
  // A shipping mark identifies a lot only for as long as it belongs to one. It covers a
  // whole contract, so the 2608174 CFS carries 1324003000A on both of its halves - and its
  // second DM row was written before the number was issued ("DM No. 13-DM-26-"), leaving
  // that half with nothing but the mark. Matched on the mark, a shipment of lifts 09-12
  // could be handed lifts 19-22's sixty-four case numbers. So where a mark stood in for a
  // missing DM number and another lot on the sheet carries the same mark, the lift number
  // takes over as the reference and the mark drops to a secondary one.
  // An escalator sheet names its Delivery Memo and then breaks the job into "A1 (E1)" and
  // "A2 (E2)", each with its own figures and cases. The DM row has already pushed a lot by
  // then, which stays empty and reads as a lot that arrived with nothing. Where a block
  // holds other lots that do carry figures or cases, an empty one is the heading they sit
  // under, and is dropped.
  for (const b of blocks) {
    const carries = (l) => l.pkgs || l.kg || l.cbm || (l.caseCodes || []).length || (l.caseNumbers || []).length;
    if (b.lots.some(carries)) b.lots = b.lots.filter(carries);
  }
  const allLots = blocks.flatMap((b) => b.lots);
  const refUse = new Map();
  for (const l of allLots) {
    for (const ref of [l.lotRef, l.altRef].filter(Boolean)) {
      refUse.set(ref, (refUse.get(ref) || 0) + 1);
    }
  }
  for (const l of allLots) {
    if (l.markOnly && l.unitCode && (refUse.get(l.lotRef) || 0) > 1) {
      l.altRef = l.lotRef;
      l.lotRef = l.unitCode;
    }
    delete l.markOnly;
  }
  return { blocks, orphanFigureRows, ownDmRef };
}
function guessFieldsFromWorkbook(wb, opts) {
  const sheetIndex = 0;
  const sheet = wb.Sheets[wb.SheetNames[sheetIndex]];
  let rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  // The same grid unformatted. A cell holding a real date comes back as a Date here, while
  // the formatted grid above renders it through the cell's own format - which on these
  // sheets is often month-first, "6/13/26". A date somebody typed as text stays text in
  // both. That difference is the only dependable way to tell "07/01/2026" meaning the 7th
  // of January from a formatted cell meaning the 1st of July, and getting it wrong is
  // invisible for the twelve days a month where both readings work.
  let rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true, cellDates: true });
  const printArea = getPrintArea(wb, sheetIndex);
  if (printArea) {
    if (printArea.lastRow) rows = rows.slice(0, printArea.lastRow);
    if (printArea.lastCol) rows = rows.map((r) => (r || []).slice(0, printArea.lastCol));
  }
  const flatText = rows.map((r) => r.join(" ")).join("\n");

  // The site is normally where the job sheet is going TO. A Return runs the other way: the
  // goods come from the site back to the yard, so the sheet's TO is the depot - "石崗(一)倉" -
  // and reading it as the project files the whole job under the storage yard. The site of a
  // Return is its FROM. Which end to read is decided by the sheet itself rather than by the
  // filename, since a document that arrives without one still has to land in the right place.
  // Nothing on the page says "return" - these sheets are all headed 工單 / JOB SHEET - so
  // the caller has to say which kind it is. It knows, from the filename.
  const goingBack = String((opts && opts.docType) || "").toLowerCase() === "return"
    || /\bRETURN\b|\u9000\u56de|\u9000\u5009/i.test(rows.slice(0, 24).map((r) => r.join(" ")).join(" "));
  const siteBlock = findAddressLines(rows, goingBack ? JOBSHEET_LABEL_ALIASES.from : JOBSHEET_LABEL_ALIASES.to);

  const out = {
    client: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.account),
    jobNumber: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.jobNo),
    date: "",
    projectEn: siteBlock.en,
    projectZh: siteBlock.zh,
    orderedBy: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.orderedBy),
    poNumber: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.poNo),
    jobRef: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.jobRef),
    ssDoNo: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.ssDoNo),
    unitCode: "",
    referJobNumber: "",
    referDate: "",
    packageCount: "",
    weightKg: "",
    volumeCbm: "",
  };

  const rawDate = findLabelValue(rows, JOBSHEET_LABEL_ALIASES.date);
  // The same cell unformatted. When Excel holds it as a real date this is a Date object and
  // there is nothing to interpret; when it is text, this is the text and the day-first
  // reading applies.
  const rawDateCell = findLabelValue(rawRows, JOBSHEET_LABEL_ALIASES.date);
  if (rawDate) {
    // A DATE cell Excel holds as a real date arrives here already formatted month-first
    // ("6/5/26"), so Date() reads it correctly. A hand-typed one is text, written day-first
    // and punctuated however it was typed - "29-/5/2026" on the 2605199 CFS sheet - which
    // Date() rejects outright, leaving the row with no date at all.
    // A date typed by hand rather than held as a real date shows it: "11- /1/2025",
    // "29-/5/2026" - a separator struck twice where the writer corrected themselves. Those
    // are written day-first, and Date() reads "11- /1/2025" as the 1st of November without
    // complaining, so they must not reach it.
    // A span reads as hand-typed too, so Date() never gets a chance to make "10-13/5/2026"
    // into something plausible-looking and wrong.
    const handTyped = /[\-\/.]\s*[\-\/.]/.test(String(rawDate))
      || /^\s*\d{1,2}\s*-\s*\d{0,2}\s*[\/\.\-]/.test(String(rawDate));
    // Anything written as d/m/y is read day-first, always. These sheets are Hong Kong
    // sheets and 07/01/2026 is the 7th of January, but Date() reads it American-style as
    // the 1st of July - silently, and only for the twelve days of each month where both
    // readings are possible, which is exactly when nobody notices. Return 1 came through
    // six months late that way. A real date cell is unambiguous and still goes to Date().
    // The unformatted cell first, since a real date there settles it outright; otherwise the
    // formatted text, read day-first.
    out.date = parseHKDate(rawDateCell instanceof Date ? rawDateCell : rawDate);
  }

  const referMatches = [...flatText.matchAll(/ref(?:er)?\.?\s*(?:to\s+)?job\s*no\.?\s*([A-Za-z0-9\-]+)\s*(?:on\s*([\d\/\.\- ]+\d))?/gi)];
  if (referMatches.length) {
    out.referJobNumber = [...new Set(referMatches.map((m) => m[1].trim()))].join(", ");
    const firstWithDate = referMatches.find((m) => m[2]);
    if (firstWithDate) out.referDate = parseSheetDayFirstDate(firstWithDate[2]);
  }

  // The per-arrival breakdown of the sheet - see parseJobSheetBlocks.
  const { blocks, orphanFigureRows, ownDmRef } = parseJobSheetBlocks(rows);
  out.refBlocks = blocks.map((b) => ({
    refJobNumber: b.refJobNumber,
    refDate: parseSheetDayFirstDate(b.refDateRaw),
    shkNumber: b.shkNumber,
    liftNo: b.liftNo,
    lots: b.lots,
    pkgs: b.pkgs, kg: b.kg, cbm: b.cbm,
    caseNumbers: [...new Set(b.lots.flatMap((l) => l.caseNumbers || []))].sort((x, y) => x - y),
  }));
  // Case markings keyed both ways, because an arrival can be found either by the lot it
  // holds (L52) or by the job number it came in under (2605199), depending on which of
  // the two the delivery sheet managed to name.
  out.caseMarksByLot = {};
  out.caseMarksByRef = {};
  // Cases written as full markings rather than numbers, keyed by the DM number, the
  // shipping mark and the lift - whichever of the three the packing list was filed under.
  out.caseCodesByLot = {};
  // Where a lot states a package count and also lists its cases, the two should agree. A
  // disagreement means the list was mis-read or the sheet is short, and either way it is
  // worth saying so rather than quietly working from the shorter number.
  out.caseCountMismatches = [];
  // Flagged so a lot with no cases can say whether the sheet is silent or whether its list
  // was pasted in as a picture, which nothing can read out of the cells.
  out.hasPastedContentImage = sheetHasPastedContentImage(wb);
  // A reference that more than one lot answers to cannot say which of them a case list
  // belongs to. Merging the two lists under it, as this used to, hands every lot on the
  // sheet the whole sheet's cases - the contract number 1324003000A sits on both halves of
  // the 2608174 CFS, 128 cases between them. Such a key is dropped instead, leaving the
  // references that do identify one lot to do the matching.
  const keysFor = (lot) => [lot.lotRef, lot.altRef, lot.unitCode].filter(Boolean).map((k) => String(k).toUpperCase());
  const keyOwners = new Map();
  for (const b of out.refBlocks) {
    for (const lot of b.lots) {
      for (const k of new Set(keysFor(lot))) keyOwners.set(k, (keyOwners.get(k) || 0) + 1);
    }
  }
  const ambiguousKey = (k) => (keyOwners.get(k) || 0) > 1;
  for (const b of out.refBlocks) {
    for (const lot of b.lots) {
      if ((lot.caseCodes || []).length) {
        const stated = Number(lot.pkgs);
        if (stated > 0 && stated !== lot.caseCodes.length) {
          out.caseCountMismatches.push({ lot: lot.lotRef || lot.unitCode || "", stated, listed: lot.caseCodes.length });
        }
        for (const key of [lot.lotRef, lot.altRef, lot.unitCode]) {
          if (!key) continue;
          const k = String(key).toUpperCase();
          if (ambiguousKey(k)) continue;
          out.caseCodesByLot[k] = {
            codes: [...new Set([...((out.caseCodesByLot[k] || {}).codes || []), ...lot.caseCodes])],
            text: lot.caseText || "",
          };
        }
      }
      if (!(lot.caseNumbers || []).length) continue;
      for (const key of [lot.unitCode, lot.lotRef, `${lot.lotRef}/${lot.unitCode}`]) {
        if (!key) continue;
        const k = String(key).toUpperCase();
        if (ambiguousKey(k)) continue;
        const prev = out.caseMarksByLot[k];
        out.caseMarksByLot[k] = {
          numbers: [...new Set([...(prev ? prev.numbers : []), ...lot.caseNumbers])].sort((x, y) => x - y),
          text: [prev && prev.text, lot.caseText].filter(Boolean).join(" "),
          // "#3,7/19" says these are cases of a 19-case lot. Kept so a case can be matched
          // on its full code rather than its number alone - 3 of 19 is not 3 of 20.
          lotCases: prev && prev.lotCases && prev.lotCases !== lot.lotCases ? null : lot.lotCases,
        };
      }
    }
    // Keyed by job number only where the block names no lot at all. A job number covers
    // every lot that arrived under it - 2605199 brought in L52, L53 and four L32 lots - so
    // a block that does name its lot must be matched on that lot, or its case numbers get
    // applied to every other lot on the same job.
    const named = b.lots.some((l) => l.unitCode || l.lotRef);
    if (b.refJobNumber && b.caseNumbers.length && !named) {
      const sizes = [...new Set(b.lots.map((l) => l.lotCases).filter((n) => n != null))];
      out.caseMarksByRef[b.refJobNumber] = {
        numbers: b.caseNumbers,
        text: b.lots.map((l) => l.caseText).filter(Boolean).join(" "),
        lotCases: sizes.length === 1 ? sizes[0] : null,
      };
    }
  }
  // A sheet closing one arrival names one SHK reference and it belongs on the row; a sheet
  // closing several names one each, and picking either would be wrong, so the field is
  // left for the user rather than guessed at.
  const shkNumbers = [...new Set(out.refBlocks
    .flatMap((b) => (b.lots.length ? b.lots.map((l) => l.shkNumber || b.shkNumber) : [b.shkNumber]))
    .filter(Boolean))];
  out.shkNumber = shkNumbers.length === 1 ? shkNumbers[0] : "";
  // The Delivery Memo a job went out under - "Refer to 13-DM-26-0228" at the foot of the
  // sheet - is the client's own reference for that delivery, so it fills the field where
  // the sheet offered nothing else for it. Set after the SHK references are gathered, or
  // it would be written and then overwritten.
  if (!out.shkNumber && ownDmRef) out.shkNumber = ownDmRef;

  const liftMatch = flatText.match(/lift\s*no\.?\s*(?:phase\s*\d+\s*)?(#[0-9,\s]+)/i)
    || flatText.match(/unit\s*(?:no\.?|code)\s*[:\-]?\s*([A-Za-z0-9\-\.\/#, ]{2,30})/i)
    || flatText.match(/escalator\s*no\.?\s*([A-Za-z0-9 &,]{2,20})/i);
  if (liftMatch) {
    out.unitCode = liftMatch[1].trim().replace(/\s+/g, " ");
  } else {
    // Fallback: some sheets list unit codes only as "REFNUMBER/E1", "REFNUMBER2/E2" lines
    // rather than a labeled heading - collect every such pair's unit code.
    const pairMatches = [...flatText.matchAll(/\b\d{6,12}\/([A-Za-z]\d+)\b/g)];
    if (pairMatches.length) out.unitCode = [...new Set(pairMatches.map((m) => m[1]))].join(", ");
  }

  // Prefer the explicit "共:" / "Total:" line for aggregate totals - a job sheet with
  // multiple case groups before the total would otherwise match the FIRST group's
  // numbers instead of the actual total.
  // Take the LAST such line that actually carries figures. The 2501043 Devan opens with a
  // note reading "(共:3 個地盤)" - three sites - which is a count of sites, not a total of
  // anything, and matching the first 共 on the page read that as the sheet's totals and
  // then fell through to the first case group's 7 PKGS / 4,233 KGS.
  const sheetTotalLines = [...flatText.matchAll(/(?:^|\n)([^\n]*(?:\u5171|total)[:\uff1a][^\n]*)/gi)]
    .map((m) => m[1])
    .filter((line) => /\d/.test(line) && /(PKGS?|KGS?|CBM)/i.test(line));
  const totalsText = sheetTotalLines.length ? sheetTotalLines[sheetTotalLines.length - 1] : flatText;
  const pkgsMatch = totalsText.match(/(\d+)\s*PKGS?/i) || flatText.match(/(\d+)\s*PKGS?/i);
  if (pkgsMatch) out.packageCount = pkgsMatch[1];
  const kgsMatch = totalsText.match(/([\d,]+(?:\.\d+)?)\s*KGS?/i) || flatText.match(/([\d,]+(?:\.\d+)?)\s*KGS?/i);
  if (kgsMatch) out.weightKg = kgsMatch[1].replace(/,/g, "");
  const cbmMatch = totalsText.match(/([\d,]+(?:\.\d+)?)\s*CBM/i) || flatText.match(/([\d,]+(?:\.\d+)?)\s*CBM/i);
  if (cbmMatch) out.volumeCbm = cbmMatch[1].replace(/,/g, "");

  // One sheet usually covers several lots, each closed by its own "共:" line - the ES1
  // Devan sheet has one for the escalator and two more for elevator lots below it.
  // Collect every total line together with the heading text directly above it, so each
  // lot's declared figures can be matched to the Incoming shipment they actually belong
  // to instead of assuming the first total on the page covers the whole file.
  out.declaredTotalsList = [];
  const totalLines = flatText.split("\n");
  for (let li = 0; li < totalLines.length; li++) {
    if (!/(?:共|total)\s*[:\uff1a]/i.test(totalLines[li])) continue;
    // "OVERSIZE TOTAL: 29.68 CBM" and "NORMAL SIZE TOTAL:" break a block down by case
    // type - they are not a lot's declared total, and treating them as one lets them
    // claim lots that belong to the block's real total line further down.
    if (/(?:over\s*size|oversize|normal\s*size)\s*total/i.test(totalLines[li])) continue;
    const pk = totalLines[li].match(/(\d+)\s*PKGS?/i);
    const kg = totalLines[li].match(/([\d,]+(?:\.\d+)?)\s*KGS?/i);
    const cb = totalLines[li].match(/([\d,]+(?:\.\d+)?)\s*CBM/i);
    if (!pk && !kg && !cb) continue;
    // Walk back to the previous total line, collecting every heading in between. A block
    // can list many lots before closing them with one figure - this sheet's second block
    // names seven (L8, L3, L6, L4, L12, L1 and more) before "共: 203 PKGS" - so the scan
    // has to run to the block boundary rather than stop after the first few, or the lots
    // furthest from the total would never be tied to it.
    const heading = [];
    for (let bi = li - 1; bi >= 0 && heading.length < 60; bi--) {
      const prev = String(totalLines[bi] || "").trim();
      if (!prev) continue;
      if (/^C\/S\s*NO\.?/i.test(prev)) continue; // case rows carry no lot identity
      // Skip the same oversize/normal-size sub-totals on the way back: they sit inside a
      // block, so treating one as the block boundary would orphan every lot above it.
      if (/(?:over\s*size|oversize|normal\s*size)\s*total/i.test(prev)) continue;
      if (/(?:共|total)\s*[:\uff1a]/i.test(prev)) break; // reached the block above
      heading.push(prev);
    }
    out.declaredTotalsList.push({
      pkgs: pk ? pk[1] : "",
      kg: kg ? kg[1].replace(/,/g, "") : "",
      cbm: cb ? cb[1].replace(/,/g, "") : "",
      context: heading.reverse().join(" ").slice(0, 2000),
    });
  }

  // A sheet that closes several lots with one "共:" line forces that figure to be shared
  // out pro-rata, which invents numbers it never stated: this delivery declares 11.44 cbm
  // against 2605126 and 6.80 against 2605199, but splitting the 18.24 total over four
  // cases each would put 9.12 on both. Every lot's own C/S NO. line already carries the
  // real figure, so those are used in front of the grand total - but only once they have
  // been checked back against it, because a sheet whose C/S rows are per-case rather than
  // per-lot would otherwise have a single case's weight read as a whole lot's.
  // What the sheet states its figures for. A lot that carries its own weight or volume -
  // whether trailing on its own line or on the C/S NO. row beneath it - is a unit in its
  // own right; a block that states one figure over several lots is a single unit covering
  // all of them. Each unit keeps the job and SHK numbers it was stated under, so the
  // figure can be labelled with them rather than only by lot.
  const declaredUnits = [];
  for (const b of out.refBlocks) {
    const lotsWithFigures = b.lots.filter((l) => l.kg || l.cbm);
    if (lotsWithFigures.length) {
      for (const l of lotsWithFigures) {
        declaredUnits.push({
          pkgs: l.pkgs || "", kg: l.kg || "", cbm: l.cbm || "",
          refJobNumber: b.refJobNumber, shkNumber: l.shkNumber || b.shkNumber,
          context: [b.refJobNumber, l.shkNumber || b.shkNumber, b.liftNo, `${l.lotRef}/${l.unitCode}`].filter(Boolean).join(" "),
        });
      }
      continue;
    }
    if ((b.lots.length || b.refJobNumber) && (b.kg || b.cbm)) {
      declaredUnits.push({
        pkgs: b.pkgs, kg: b.kg, cbm: b.cbm,
        refJobNumber: b.refJobNumber, shkNumber: b.shkNumber,
        context: [b.refJobNumber, b.shkNumber, b.liftNo, ...b.lots.map((l) => `${l.lotRef}/${l.unitCode}`)]
          .filter(Boolean).join(" "),
      });
    }
  }
  // Trust those figures only when every weight the sheet stated found a lot to belong to.
  // A sheet whose C/S NO. rows are per-case rather than per-lot leaves weights stranded
  // once the first has claimed the lot, and those strays are the signal to fall back to
  // the sheet's own "共:" total instead of reading a single case as a whole lot.
  //
  // Reconciling against that total is not the test: this delivery states 9,063.74 kg in
  // its 共: line, which covers only the three lots of its first referral and leaves the
  // 5,825.5 kg of the second out entirely. Requiring the two to agree would have thrown
  // away every per-lot figure on the page.
  if (declaredUnits.length && orphanFigureRows === 0) {
    out.declaredTotalsList = declaredUnits;
  }

  const ssShipMatch = flatText.match(/ex\s*ss\.?\s*"[^"]+"[^\n]*/i);
  if (ssShipMatch && !out.ssDoNo) out.ssDoNo = ssShipMatch[0].trim();

  // Some Devan/CFS sheets declare oversize lots explicitly, e.g.
  // "OVERSIZE CASES: L13@4.49CBM, L14@4.49CBM" - these lots have no per-case CBM data of
  // their own, so this figure is the only CBM basis available for them, and oversize
  // cargo bills at a higher rate.
  // OVERSIZE CASES, as Irene writes them:
  //   P2#14/21@3.92CBM, P4#7/29@3.96CBM, #8/29@3.26CBM, #21/29@5.28CBM
  //   L13@4.49CBM, L14@4.49CBM
  // A lot can carry several oversize cases, and after the first the lot prefix is dropped
  // - "#8/29" still belongs to P4. Some sheets name no case at all and give one figure for
  // the whole lot. All three shapes land in the same structure: lot -> [{code, cbm}].
  out.oversizeByLot = {};
  const oversizeSectionMatch = flatText.match(/oversize\s*cases?[:\uff1a]?\s*([\s\S]{0,600}?)(?:\n\s*\n|$)/i);
  if (oversizeSectionMatch) {
    const entries = [...oversizeSectionMatch[1].matchAll(
      /(?:([A-Za-z][A-Za-z0-9\-]*)\s*)?(?:#\s*([0-9]+(?:\s*\/\s*[0-9]+)?))?\s*@\s*([\d,]+(?:\.\d+)?)\s*CBM/gi
    )];
    let carriedLot = "";
    for (const e of entries) {
      const lot = (e[1] || "").trim() || carriedLot;
      if (!lot) continue;
      carriedLot = lot;
      const code = (e[2] || "").replace(/\s+/g, "");
      const cbm = e[3].replace(/,/g, "");
      if (!out.oversizeByLot[lot]) out.oversizeByLot[lot] = [];
      out.oversizeByLot[lot].push({ code, cbm });
    }
  }


  return out;
}

// Reads a C/S NO. marking off a job sheet into case codes. The notation varies even
// within one sheet - "1-12/23", "1,2,3/3", "1/2, 2/2", "1,3,4,5/5" - and the lot size may
// be repeated after every case rather than stated once at the end. A trailing note in
// brackets ("1/1 (17/F)") is a floor reference, not part of the numbering.
function parseCaseSpec(spec) {
  // The marking is stored with the "#" it was written with, so take that off first.
  let s = String(spec || "").replace(/\([^)]*\)/g, " ").replace(/#/g, " ").trim();
  if (!s) return { codes: [], numbers: [], lotCases: null };
  const totals = [...s.matchAll(/\/\s*(\d+)/g)].map((m) => Number(m[1]));
  const lotCases = totals.length ? totals[totals.length - 1] : null;
  s = s.replace(/\/\s*\d+/g, " ").replace(/\s+/g, " ").trim();
  if (!/^[\d\s,\-]+$/.test(s)) return { codes: [], numbers: [], lotCases };
  const numbers = [...parseRangeInput(s)].sort((a, b) => a - b);
  const codes = numbers.map((n) => (lotCases ? `${n}/${lotCases}` : String(n)));
  return { codes, numbers, lotCases };
}
function siteKeyFor(en, zh) {
  return [(en || "").trim().toLowerCase(), (zh || "").trim()].filter(Boolean);
}
function normalizeSiteForMatch(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/^(site at|no\.?)\s*/i, "")
    .replace(/,?\s*(hong ?kong|hk|wan ?chai)\s*$/i, "")
    .replace(/[^\w\u4e00-\u9fff]+/g, " ")
    .trim();
}
function sitesLooselyMatch(enA, zhA, enB, zhB) {
  const aEn = normalizeSiteForMatch(enA), bEn = normalizeSiteForMatch(enB);
  const aZh = String(zhA || "").trim(), bZh = String(zhB || "").trim();
  if (aEn && bEn && (aEn.includes(bEn) || bEn.includes(aEn))) return true;
  if (aZh && bZh && (aZh.includes(bZh) || bZh.includes(aZh))) return true;
  return false;
}
// Job numbers are the same number however they were typed - "2605199", "2605199 " and
// "26-05199" all refer to one arrival - so they are compared on digits and letters alone.
function jobNosMatch(a, b) {
  const na = String(a || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const nb = String(b || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  return !!na && na === nb;
}
// A delivery sheet may name a lot more loosely than the arrival record does - "L32" for
// what was checked in as "L32-01". Matching allows one to be a prefix of the other only
// where the next character is a separator, so "L32" finds "L32-01" while "L3" does not.
function lotTokenMatches(a, b) {
  const A = String(a || "").toUpperCase().replace(/\s+/g, "");
  const B = String(b || "").toUpperCase().replace(/\s+/g, "");
  if (!A || !B) return false;
  for (const partA of A.split(/[,;]/).filter(Boolean)) {
    for (const partB of B.split(/[,;]/).filter(Boolean)) {
      if (partA === partB) return true;
      const [s, l] = partA.length <= partB.length ? [partA, partB] : [partB, partA];
      if (s.length >= 2 && l.startsWith(s) && /[^A-Z0-9]/.test(l.charAt(s.length))) return true;
    }
  }
  return false;
}
// Parses typed case-number input like "1,3-5,7" into a Set of integers, supporting
// both single numbers and ranges.
function parseRangeInput(text) {
  const nums = new Set();
  for (const part of String(text || "").split(",").map((s) => s.trim()).filter(Boolean)) {
    const rangeMatch = part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const a = Number(rangeMatch[1]), b = Number(rangeMatch[2]);
      for (let n = Math.min(a, b); n <= Math.max(a, b); n++) nums.add(n);
    } else if (/^\d+$/.test(part)) {
      nums.add(Number(part));
    }
  }
  return nums;
}
// Extracts the leading case number from a code like "26/40" -> 26, so typed numbers can
// be matched against it regardless of the "/total" suffix.
function codeLeadingNumber(code) {
  const m = String(code).match(/^\d+/);
  return m ? Number(m[0]) : null;
}
// A case marking is not always unique within a shipment - Mitsubishi's 13-DM-26-0500 has
// two lines marked 01C01, and both are genuine. So the selection is a multiset: a marking
// chosen once accounts for one case, not for every case bearing that marking. Counting by
// set membership made 26 selected cases weigh in as 27 packages.
// Which occurrence of a repeated marking a chip stands for. With two cases both marked
// 01C01, "is this one selected?" cannot be answered by the code alone: the nth case
// bearing a marking is selected when the marking was chosen at least n times.
function nthOccurrence(packages, index) {
  const code = (packages[index] || {}).code;
  let n = 0;
  for (let i = 0; i < index; i++) if (packages[i].code === code) n += 1;
  return n;
}
function occurrenceSelected(packages, index, codes) {
  const code = (packages[index] || {}).code;
  const chosen = (codes || []).filter((c) => c === code).length;
  return nthOccurrence(packages, index) < chosen;
}
function toggleOccurrence(packages, index, codes) {
  const code = (packages[index] || {}).code;
  const cur = codes || [];
  if (occurrenceSelected(packages, index, cur)) {
    const out = [...cur];
    out.splice(out.indexOf(code), 1);
    return out;
  }
  return [...cur, code];
}
function sumSelectedPackages(packages, codes) {
  const wanted = new Map();
  for (const c of codes || []) wanted.set(c, (wanted.get(c) || 0) + 1);
  const matched = [];
  for (const p of packages || []) {
    const left = wanted.get(p.code) || 0;
    if (left <= 0) continue;
    wanted.set(p.code, left - 1);
    matched.push(p);
  }
  return {
    count: matched.length,
    weight: matched.reduce((s, p) => s + (Number(p.weightKg) || 0), 0),
    cbm: matched.reduce((s, p) => s + (Number(p.cbm) || 0), 0),
  };
}
// ---------------------------------------------------------------------------
// Declared totals (Devan/CFS sheet) vs packing-list totals.
//
// Two weights legitimately exist for the same cargo. The manufacturer's detail
// packing list gives a gross weight per case; the Devan/CFS sheet gives a total taken
// from the carrier paperwork, which on escalator jobs runs heavier because it includes
// the crating and skids on the truss sections. Worked example - SHK0959/25 ES1: the
// four parts boxes agree exactly (1,130 kg both ways) while 501/502/503 come in
// +380/+635/+375, a 1,390 kg gap on an otherwise identical shipment.
//
// Per Schindler (Irene), the Devan figure is the accurate one and any real discrepancy
// is settled on paperwork that never reaches this server. So declared wins for weight,
// volume and billing, the packing-list sum is kept alongside it as the per-case record,
// and the difference is shown rather than treated as an error.
// ---------------------------------------------------------------------------
// Older rows stored one figure per lot; keep them readable alongside the case list.
function normaliseOversize(v) {
  if (!v) return { checked: false, cases: [{ code: "", cbm: "" }] };
  if (Array.isArray(v.cases)) return { checked: !!v.checked, cases: v.cases.length ? v.cases : [{ code: "", cbm: "" }] };
  return { checked: !!v.checked, cases: [{ code: typeof v.cases === "string" ? v.cases : "", cbm: v.cbm || "" }] };
}
function cleanOversizeCases(cases) {
  return (cases || [])
    .map((c) => ({ code: String(c.code || "").trim(), cbm: String(c.cbm || "").trim() }))
    .filter((c) => Number(c.cbm) > 0);
}
function oversizeCasesOf(item) {
  const list = cleanOversizeCases(item.oversizeCases);
  if (list.length) return list;
  // Pre-existing items carry a single lot-level figure and no case codes.
  if (item.isOversize && Number(item.oversizeCbm) > 0) return [{ code: "", cbm: String(item.oversizeCbm) }];
  return [];
}
function oversizeCbmTotal(cases) {
  return Math.round(cleanOversizeCases(cases).reduce((s, c) => s + Number(c.cbm), 0) * 1000) / 1000;
}
// Codes that carry their own oversize CBM, and therefore their own rate tier.
function oversizeCaseMap(item) {
  const m = new Map();
  for (const c of oversizeCasesOf(item)) if (c.code) m.set(c.code, Number(c.cbm));
  return m;
}
function declaredContextKey(s) {
  return String(s || "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
}
// Lot identifiers arrive in inconsistent shapes: "ES-1" against a unit code of "ES1",
// "60726103/L7" against "L7". Comparing on a fully stripped string is unsafe - once the
// separators go, "60726103/L7 60726117/L8" reads as "...L760726117L8", and a boundary
// test for L7 sees a digit after it and fails. So split the heading into alphanumeric
// tokens, add the joins of adjacent tokens (which recovers "ES" + "1" as "ES1"), and
// require the identifier to equal one of those outright. Exact equality also keeps "ES1"
// from matching "ES12" without needing a boundary rule at all.
function declaredContextTokens(context) {
  const tokens = String(context || "").toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
  const keys = new Set(tokens);
  for (let i = 0; i < tokens.length - 1; i++) {
    keys.add(tokens[i] + tokens[i + 1]);
    if (i < tokens.length - 2) keys.add(tokens[i] + tokens[i + 1] + tokens[i + 2]);
  }
  return keys;
}
function lotIdentityMatches(context, lot) {
  return lotUnitMatches(context, lot) || lotShkMatches(context, lot);
}
// The lot code is what a total line names to say which lot it is for. It is matched on its
// own, ahead of the SHK reference, because one SHK covers a whole consignment: SHK0395/26
// spans L52 and L53, so a line naming "60766022/L53" would otherwise be claimed by every
// lot filed under that reference - which is how a total stated for L53 ended up covering
// L34-S1 and L34-S2 instead.
function lotUnitMatches(context, lot) {
  const keys = declaredContextTokens(context);
  if (!keys.size) return false;
  for (const part of String(lot.unitCode || "").split(/[,;]/)) {
    const needle = declaredContextKey(part);
    if (needle.length >= 2 && keys.has(needle)) return true;
  }
  return false;
}
function lotShkMatches(context, lot) {
  const keys = declaredContextTokens(context);
  if (!keys.size) return false;
  for (const part of String(lot.shkNumber || "").split(/[,;]/)) {
    const needle = declaredContextKey(part);
    if (needle.length >= 2 && keys.has(needle)) return true;
  }
  return false;
}
function declaredNum(v) {
  const n = Number(String(v == null ? "" : v).replace(/,/g, ""));
  return Number.isFinite(n) && n > 0 ? n : null;
}
// Groups the matched lots by the total line that covers them. A sheet may close each lot
// separately (the ES1 Devan) or close several at once (the L7/L8/L9 delivery, whose only
// total covers all twelve packages); the group is what the sheet actually states, and is
// therefore the thing the user edits.
// The order number identifies a consignment exactly where a lift code does not. Two
// separate Incoming shipments can both be L32-01 - a part-load followed by the rest - so a
// total the sheet states for "60759188/L32-01" belongs to the one carrying order 60759188
// alone. Matching on the lift code alone claimed both, which turned a figure the sheet
// stated for a single lot into a shared one, and a shared figure is treated as an estimate
// that loses to real per-case weights: the 16,527 kg on the 2605126 CFS sheet was being
// shown as a "share" and then quietly not used at all.
function lotOrderNumbers(lot) {
  const nums = new Set();
  for (const p of lot.packages || []) {
    const n = String(p.orderNo || "").trim();
    if (/^\d{6,12}$/.test(n)) nums.add(n);
  }
  return nums;
}
function lotOrderMatches(context, lot) {
  const keys = declaredContextTokens(context);
  for (const n of lotOrderNumbers(lot)) if (keys.has(n)) return true;
  return false;
}
function groupDeclaredLots(list, lots) {
  const all = list || [];
  const groups = [];
  const claimed = new Set();
  const unnamed = [];
  for (const line of all) {
    const available = lots.filter((lot) => !claimed.has(lot.id));
    // Strongest identifier first. An order number pins a line to one consignment; a lot
    // code pins it to one lot; an SHK reference only narrows it to a whole consignment and
    // is a last resort for sheets that name nothing else.
    const group = (() => {
      for (const test of [lotOrderMatches, lotUnitMatches, lotShkMatches]) {
        const hit = available.filter((lot) => test(line.context, lot));
        if (hit.length) return hit;
      }
      return [];
    })();
    if (group.length) {
      group.forEach((lot) => claimed.add(lot.id));
      groups.push({ line, lots: group });
    } else unnamed.push(line);
  }
  // Nothing on the sheet named a lot: fall back to its sole total covering everything.
  if (claimed.size === 0 && unnamed.length === 1 && lots.length) groups.push({ line: unnamed[0], lots: [...lots] });
  return groups;
}
function declaredGroupKey(group) {
  return group.lots.map((l) => l.id).sort().join("|");
}
// Shares one declared total out over the lots it covers, pro-rata on the packing-list
// figures of the cases selected - the only per-case data there is.
function splitDeclaredAcrossLots(total, lots, listedFor) {
  const out = {};
  if (!total) return out;
  const totPkgs = declaredNum(total.pkgs);
  const totKg = declaredNum(total.kg);
  const totCbm = declaredNum(total.cbm);
  const listed = lots.map((lot) => listedFor(lot) || { count: 0, weight: 0, cbm: 0 });
  const shareOf = (key, i) => {
    const denom = listed.reduce((s, l) => s + (Number(l[key]) || 0), 0);
    if (denom > 0) return (Number(listed[i][key]) || 0) / denom;
    // Elevator lots often carry no per-case weight or volume at all. Case count is the
    // next best proxy - a 32-case lot should take more of the total than an 18-case one.
    const byCount = listed.reduce((s, l) => s + (Number(l.count) || 0), 0);
    if (byCount > 0) return (Number(listed[i].count) || 0) / byCount;
    return lots.length ? 1 / lots.length : 0;
  };
  const single = lots.length === 1;
  // Rounding each share independently leaves the split a little short of or over the
  // stated total. The last lot takes the remainder so the parts always add back to the
  // figure on the sheet - a depot total that doesn't reconcile invites a recount.
  const spread = (total, key, decimals) => {
    if (total == null) return lots.map(() => "");
    if (single) return [String(total)];
    const f = 10 ** decimals;
    const vals = [];
    let used = 0;
    lots.forEach((lot, i) => {
      if (i === lots.length - 1) {
        vals.push(Math.round((total - used) * f) / f);
      } else {
        const v = Math.round(total * shareOf(key, i) * f) / f;
        used += v;
        vals.push(v);
      }
    });
    return vals.map(String);
  };
  const kgs = spread(totKg, "weight", 1);
  const cbms = spread(totCbm, "cbm", 3);
  lots.forEach((lot, i) => {
    out[lot.id] = {
      // A per-lot package count is only meaningful when the line describes one lot.
      // Splitting it would invent a number and fire a false package-count warning.
      pkgs: single && totPkgs != null ? String(totPkgs) : "",
      kg: kgs[i],
      cbm: cbms[i],
      split: !single,
    };
  });
  return out;
}
// `overridesByGroup` holds totals the user has corrected, keyed by declaredGroupKey.
function distributeDeclaredAcrossLots(list, lots, listedFor, overridesByGroup) {
  const out = {};
  if (!lots || !lots.length) return out;
  for (const group of groupDeclaredLots(list, lots)) {
    const key = declaredGroupKey(group);
    const total = (overridesByGroup && overridesByGroup[key]) || group.line;
    Object.assign(out, splitDeclaredAcrossLots(total, group.lots, listedFor));
  }
  return out;
}
// `totals` is the packing-list sum of the selected cases, `declared` the sheet figures.
function computeDeclaredVariance(totals, declared) {
  if (!declared) return null;
  const out = { any: false };
  const pair = (key, listed, raw) => {
    const d = declaredNum(raw);
    if (d == null || !(listed > 0)) return null;
    const delta = d - listed;
    const pct = (delta / listed) * 100;
    if (Math.abs(pct) >= 0.05) out.any = true;
    return { declared: d, listed, delta, pct };
  };
  out.pkgs = pair("pkgs", totals.count, declared.pkgs);
  out.kg = pair("kg", totals.weight, declared.kg);
  out.cbm = pair("cbm", totals.cbm, declared.cbm);
  if (out.pkgs && out.pkgs.delta !== 0) out.any = true;
  return out;
}
// Recomputes an item's headline weight/volume from its arrival batches. A batch that
// carried declared figures contributes those; a batch checked in without a sheet (the
// Incoming tab, manual entry) contributes the packing-list weight of its own cases. An
// explicit oversize CBM still outranks everything, as it did before.
function recomputeItemTotals(item) {
  const pkgs = item.packages || [];
  const listedWeight = pkgs.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
  const listedCbm = pkgs.reduce((s, p) => s + (Number(p.cbm) || 0), 0);
  const arrivals = (item.arrivals || []).filter((a) => !a.voided);
  const anyDeclaredKg = arrivals.some((a) => a.declared && declaredNum(a.declared.kg) != null);
  const anyDeclaredCbm = arrivals.some((a) => a.declared && declaredNum(a.declared.cbm) != null);

  const next = {
    ...item,
    weightPackingListKg: listedWeight ? String(Math.round(listedWeight * 10) / 10) : "",
    volumeCbmPackingList: listedCbm ? String(Math.round(listedCbm * 1000) / 1000) : "",
    weightSource: anyDeclaredKg ? "declared" : "",
    volumeSource: "",
  };
  const oversizeCbm = item.isOversize ? declaredNum(item.oversizeCbm) : null;
  if (oversizeCbm != null) next.volumeSource = "oversize";
  else if (anyDeclaredCbm) next.volumeSource = "declared";
  if (!anyDeclaredKg && !anyDeclaredCbm) return next;

  // Weight and volume are resolved independently, so a sheet that states only one of
  // them doesn't drag the other off the packing list. Within each, a batch that carried
  // a sheet figure contributes it; a batch checked in without one contributes the
  // packing-list total of its own cases.
  //
  // A sheet figure only wins when the sheet actually stated it for this lot. Where it is
  // this lot's share of a total covering many lots - the Devan that ran fourteen lifts
  // together, whose kg and cbm were never divided accurately between them - it is an
  // estimate, and measured per-case weights from the packing list are the better record.
  // The lump total stays on the arrival batch either way.
  const hasListedWeights = pkgs.some((p) => Number(p.weightKg) > 0);
  const hasListedCbm = pkgs.some((p) => Number(p.cbm) > 0);
  let weight = 0;
  let cbm = 0;
  let usedKg = false, usedCbm = false, estKg = false, estCbm = false;
  for (const a of arrivals) {
    const batchListed = sumSelectedPackages(pkgs, a.codes || []);
    const split = !!(a.declared && a.declared.split);
    const dKg = a.declared ? declaredNum(a.declared.kg) : null;
    const dCbm = a.declared ? declaredNum(a.declared.cbm) : null;
    const takeKg = dKg != null && !(split && hasListedWeights);
    const takeCbm = dCbm != null && !(split && hasListedCbm);
    if (takeKg) { usedKg = true; if (split) estKg = true; }
    if (takeCbm) { usedCbm = true; if (split) estCbm = true; }
    weight += takeKg ? dKg : batchListed.weight;
    cbm += takeCbm ? dCbm : batchListed.cbm;
  }
  next.weightSource = usedKg ? (estKg ? "declared-estimated" : "declared") : "";
  if (usedCbm) next.volumeSource = estCbm ? "declared-estimated" : "declared";
  else if (next.volumeSource === "declared") next.volumeSource = "";
  if (usedKg && weight) next.weightKg = String(Math.round(weight * 10) / 10);
  // Oversize outranks the sheet's lot total on volume, and only on volume. Oversize
  // cargo is priced off its CBM alone - `oversizeCbm` picks the rate multiplier tier in
  // computeItemBillingRows and stands in as the billable volume when the cases carry no
  // CBM of their own - so a coarser lot total must not displace it. Weight has no such
  // override, which is why it takes the sheet figure unconditionally.
  if (oversizeCbm == null && usedCbm && cbm) next.volumeCbm = String(Math.round(cbm * 1000) / 1000);
  else if (oversizeCbm != null) next.volumeSource = "oversize";
  return next;
}
// Totals for a subset of an item's cases, expressed on the item's authoritative basis.
//
// sumSelectedPackages adds up the per-case packing-list figures, which is the right
// answer while cases are still on an Incoming shipment - nothing has been declared yet.
// Once cases are in inventory the item carries a Devan/CFS weight and, for oversize
// cargo, a CBM that sets the billing tier. A delivery drawn from that item has to be
// quoted on the same basis, or the depot shows one number on arrival and a different
// one on the way out. The subset's share is taken from the packing list (the only
// per-case data there is) and applied to the item's own total, falling back to a plain
// case count when the packing list carries no weights or volumes at all.
// Substitutes the sheet's declared figures for the derived ones where it states them.
function effectiveDeliveryTotals(totals, declared) {
  const kg = declared ? declaredNum(declared.kg) : null;
  const cbm = declared ? declaredNum(declared.cbm) : null;
  return {
    count: totals.count,
    weight: kg != null ? kg : totals.weight,
    cbm: cbm != null ? cbm : totals.cbm,
  };
}
function selectedItemTotals(item, codes) {
  const pkgs = item.packages || [];
  const s = sumSelectedPackages(pkgs, codes);
  const listedWeight = pkgs.reduce((acc, p) => acc + (Number(p.weightKg) || 0), 0);
  const listedCbm = pkgs.reduce((acc, p) => acc + (Number(p.cbm) || 0), 0);
  const countShare = pkgs.length ? s.count / pkgs.length : 0;
  const itemWeight = Number(item.weightKg) || 0;
  const itemCbm = Number(item.volumeCbm) || 0;
  const weight = item.weightSource === "declared" && itemWeight > 0
    ? itemWeight * (listedWeight > 0 ? s.weight / listedWeight : countShare)
    : s.weight;
  const cbm = item.volumeSource && itemCbm > 0
    ? itemCbm * (listedCbm > 0 ? s.cbm / listedCbm : countShare)
    : s.cbm;
  return { count: s.count, weight, cbm };
}
// Site pickers get long fast. Two filters keep them usable: only show sites belonging to
// the client already selected, and drop sites whose material has all left the depot -
// those are finished jobs, still on file and still searchable, just not in the way. A
// site with nothing booked against it yet is never treated as finished.
function siteMatchesItem(site, item) {
  const k = (s) => String(s || "").trim().toLowerCase();
  if (site.jobRef && item.jobRef && k(site.jobRef) === k(item.jobRef)) return true;
  if (site.siteEn && k(site.siteEn) === k(item.project)) return true;
  if (site.siteZh && k(site.siteZh) === k(item.constructionSite)) return true;
  if (site.siteEn && k(site.siteEn) === k(item.constructionSite)) return true;
  return false;
}
function siteIsCompleted(site, items) {
  const matched = (items || []).filter((i) => siteMatchesItem(site, i));
  if (matched.length === 0) return false;
  return matched.every((i) => remainingUnits(i) === 0);
}
function visibleDirectory(directory, { client, showOlder, items }) {
  return (directory || []).filter((s) => {
    if (client && client !== "Other" && s.client && s.client !== client) return false;
    if (showOlder) return true;
    return !siteIsCompleted(s, items);
  });
}
function hiddenSiteCount(directory, { client, items }) {
  return (directory || []).filter((s) => {
    if (client && client !== "Other" && s.client && s.client !== client) return false;
    return siteIsCompleted(s, items);
  }).length;
}
function groupPackagesByOrder(pkgs) {
  const hasAnyOrderNo = pkgs.some((p) => p.orderNo);
  if (!hasAnyOrderNo) return [{ orderNo: "", packages: pkgs }];
  const order = [];
  const groups = {};
  for (const p of pkgs) {
    const key = p.orderNo || "";
    if (!groups[key]) { groups[key] = []; order.push(key); }
    groups[key].push(p);
  }
  return order.map((k) => ({ orderNo: k, packages: groups[k] }));
}
// Oversize cases for one lot. A lot can have several - P4 ships #7/29, #8/29 and #21/29
// at three different volumes - and each is priced on its own CBM, so they are held as a
// list of {code, cbm} rather than one figure for the lot.
function OversizeCasesEditor({ value, onToggle, onCases, colors, t, inputClass, inputStyle }) {
  const cases = value.cases && value.cases.length ? value.cases : [{ code: "", cbm: "" }];
  const patch = (i, p) => onCases(cases.map((c, k) => (k === i ? { ...c, ...p } : c)));
  const total = cases.reduce((s, c) => s + (Number(c.cbm) || 0), 0);
  return (
    <div className="rounded px-2 py-2" style={{ background: value.checked ? colors.amberSoft : "transparent" }}>
      <label className="flex items-center gap-1.5 text-xs font-semibold" style={{ color: colors.ink }}>
        <input type="checkbox" checked={!!value.checked} onChange={(e) => onToggle(e.target.checked)} />
        {t.legacyOversizeLabel}
      </label>
      {value.checked && (
        <div className="mt-1.5 flex flex-col gap-1.5">
          {cases.map((c, i) => (
            <div key={i} className="flex flex-wrap items-center gap-2">
              <span className="text-xs" style={{ color: colors.inkFaint, width: 34 }}>{t.legacyOversizeCaseCol}</span>
              <input
                className={inputClass}
                style={{ ...inputStyle, width: 96, fontSize: 12, padding: "4px 8px" }}
                placeholder={t.legacyOversizeCasesPh}
                value={c.code || ""}
                onChange={(e) => patch(i, { code: e.target.value })}
              />
              <input
                type="number" min="0" step="0.001"
                className={inputClass}
                style={{ ...inputStyle, width: 86, fontSize: 12, padding: "4px 8px" }}
                placeholder={t.jsCbm}
                value={c.cbm || ""}
                onChange={(e) => patch(i, { cbm: e.target.value })}
              />
              {cases.length > 1 && (
                <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }}
                  onClick={() => onCases(cases.filter((_, k) => k !== i))}>{t.legacyOversizeRemove}</button>
              )}
            </div>
          ))}
          <div className="flex flex-wrap items-center gap-3">
            <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }}
              onClick={() => onCases([...cases, { code: "", cbm: "" }])}>{t.legacyOversizeAdd}</button>
            <span className="text-xs" style={{ color: colors.inkFaint }}>
              {cases.length > 1 ? t.legacyOversizeTotal(cases.length, Math.round(total * 1000) / 1000) : t.legacyOversizeHint}
            </span>
          </div>
        </div>
      )}
    </div>
  );
}
// One editable set of figures for a total the sheet states once across several lots.
function SharedDeclaredTotal({ group, value, onPatch, colors, t, inputClass, inputStyle }) {
  const names = group.lots.map((l) => l.unitCode || l.id).join(", ");
  // Name the figure by the job and SHK it was stated under, not only by the lots it lands
  // on - a sheet closing several referrals has more than one of these boxes, and the lot
  // codes alone do not say which referral each one came from.
  const ref = [
    group.line.refJobNumber ? t.legacySheetTotalJob(group.line.refJobNumber) : "",
    group.line.shkNumber || "",
  ].filter(Boolean).join(" \u00b7 ");
  return (
    <div className="mb-2 px-2 py-2 rounded" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <div className="text-xs font-semibold mb-1.5" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
        {t.legacySheetTotalLabel(names, ref)}
      </div>
      <div className="flex flex-wrap items-center gap-2">
        {[
          { key: "pkgs", ph: t.jsPkgs, w: 64 },
          { key: "kg", ph: t.jsKgs, w: 96 },
          { key: "cbm", ph: t.jsCbm, w: 88 },
        ].map((f) => (
          <input
            key={f.key}
            type="number" min="0" step="0.001"
            className={inputClass}
            style={{ ...inputStyle, width: f.w, fontSize: 12, padding: "4px 8px" }}
            placeholder={f.ph}
            value={value[f.key] || ""}
            onChange={(e) => onPatch({ [f.key]: e.target.value })}
          />
        ))}
      </div>
      <div className="text-xs mt-1" style={{ color: colors.inkFaint }}>{t.legacySheetTotalHint}</div>
    </div>
  );
}
function LegacyUploadRow({ onReplaceIncomingCases, directory, setDirectory, employees, row, siblingRows, onChange, onRemove, incoming, items, onLegacyEnrich, onAddIncoming, onProcessAll, processing, processDisabled, colors, t, lang }) {
  const inputStyle = inputStyleFor(colors);
  const set = (k) => (e) => onChange({ ...row, [k]: e.target.value });
  const itemized = JOB_SHEET_ITEMIZED.includes(row.docType);
  // A Devan or a CFS brings goods in off a packing list, so it is matched against Incoming.
  // A Return brings back cases that already went out, so what it has to find is the entry
  // they were delivered from - the same lookup a Delivery does, against inventory. Filing it
  // with the arrivals sent it hunting through packing lists it will never appear on, which
  // is why it showed none of its referral blocks.
  const canMatchIncoming = row.docType === "Devan" || row.docType === "CFS";
  const matchesInventory = row.docType === "Delivery" || row.docType === "Return";
  const matchedIncomings = canMatchIncoming && row.client && (row.projectEn || row.projectZh)
    ? (incoming || []).filter((inc) => {
        if (inc.client !== row.client) return false;
        const done = new Set(inc.checkedInCodes || []);
        const remaining = (inc.packages || []).filter((p) => !done.has(p.code));
        if (remaining.length === 0) return false;
        return sitesLooselyMatch(row.projectEn, row.projectZh, inc.project, inc.constructionSite);
      })
    : [];
  // Shipments still waiting to be checked in that this sheet did NOT match. Where any
  // exist, the client or the site is written differently on one side or the other, and
  // that is worth seeing before a second packing list gets built from the sheet on top of
  // one that is already here.
  const nearMissIncomings = (incoming || []).filter((inc) => {
    if (matchedIncomings.includes(inc)) return false;
    const done = new Set(inc.checkedInCodes || []);
    return (inc.packages || []).some((p) => !done.has(p.code));
  });
  const selectedByIncoming = row.selectedByIncoming || {};
  function setSelectedForIncoming(incId, next) {
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: next } });
  }
  function toggleIncomingCode(incId, code) {
    const cur = selectedByIncoming[incId] || [];
    const next = cur.includes(code) ? cur.filter((c) => c !== code) : [...cur, code];
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: next } });
  }
  function selectAllIncoming(incId, remainingPkgs) {
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: remainingPkgs.map((p) => p.code) } });
  }
  const oversizeByIncoming = row.oversizeByIncoming || {};
  function getOversizeFor(inc) {
    if (oversizeByIncoming[inc.id] !== undefined) return normaliseOversize(oversizeByIncoming[inc.id]);
    const detected = (row.oversizeByLot || {})[inc.unitCode];
    return detected && detected.length
      ? { checked: true, cases: detected.map((c) => ({ code: c.code || "", cbm: c.cbm || "" })) }
      : { checked: false, cases: [{ code: "", cbm: "" }] };
  }
  function setOversizeFor(inc, patch) {
    const cur = normaliseOversize(getOversizeFor(inc));
    onChange({ ...row, oversizeByIncoming: { ...oversizeByIncoming, [inc.id]: { ...cur, ...patch } } });
  }
  // Totals as declared on this sheet, per matched lot. Pre-filled from the sheet's own
  // "共:" lines and editable, because a scanned or hand-typed sheet won't always parse.
  const declaredByIncoming = row.declaredByIncoming || {};
  const incomingListedFor = (inc) => sumSelectedPackages(inc.packages, selectedByIncoming[inc.id] || []);
  // A Devan/CFS sheet names the cases arriving just as a delivery sheet names the ones
  // leaving - "C/S NO. 13-23/23" - so the same marking is used here to tick them. Where the
  // shipment does not hold one of them, saying which is missing turns "8 of 11" into
  // something answerable: on the 2605076 sheet, 21, 22 and 23 of lot 60701670/L5 are not in
  // INC-0222 at all, because that order is filed twice and they sit in the other record.
  // A shipment is easier to recognise by what the paperwork calls it than by its FS number:
  // the SHK reference it came in under, the order/commission number on its cases, and the
  // lift it is for. Shown in that order wherever a shipment is named.
  function incomingLabel(inc) {
    const orders = incomingOrders(inc);
    return [inc.shkNumber, orders.slice(0, 2).join(", "), inc.unitCode].filter(Boolean).join(" \u00b7 ") || "\u2014";
  }
  function incomingOrders(inc) {
    return [...new Set((inc.packages || []).map((p) => String(p.orderNo || "").trim()).filter(Boolean))];
  }
  function sheetMarkKeyForIncoming(inc) {
    const byLot = row.caseMarksByLot || {};
    const orders = incomingOrders(inc);
    return Object.keys(byLot).find((k) => lotTokenMatches(inc.unitCode, k) || orders.some((o) => lotTokenMatches(o, k))) || null;
  }
  function sheetMarkForIncoming(inc) {
    const key = sheetMarkKeyForIncoming(inc);
    return key ? (row.caseMarksByLot || {})[key] : null;
  }
  // One order number can legitimately appear on several packing lists - an order ships in
  // batches, and each list numbers its own cases, so "14/23" on one consignment is not the
  // same box as "14/23" on another. Where more than one matched shipment answers to the
  // same marking there is no safe way to tell which the sheet means, so nothing is ticked
  // and the choice is left to whoever has the paperwork.
  function shipmentsSharingMark(inc) {
    const key = sheetMarkKeyForIncoming(inc);
    if (!key) return [];
    return matchedIncomings.filter((other) => sheetMarkKeyForIncoming(other) === key);
  }
  function sheetSelectionForIncoming(inc) {
    const byCode = row.caseCodesByLot || {};
    const codeKeys = Object.keys(byCode);
    const codeKey = codeKeys.find((k) => lotTokenMatches(inc.unitCode, k))
      || codeKeys.find((k) => incomingOrders(inc).some((o) => lotTokenMatches(o, k)));
    if (codeKey && (byCode[codeKey].codes || []).length) {
      const done = new Set(inc.checkedInCodes || []);
      const available = (inc.packages || []).filter((p) => !done.has(p.code));
      const norm = (c) => String(c || "").toUpperCase().replace(/\s+/g, "");
      const wanted = byCode[codeKey].codes;
      // Matched through the resolver, so a sheet's shorter "03B11" still finds the packing
      // list's "03B1103" where only one case can possibly be meant.
      const resolved = new Map(wanted.map((c) => [c, resolveCaseCode(c, available)]));
      const missing = wanted.filter((c) => !resolved.get(c));
      // Same number of cases, different numbering. Which list is right is not something
      // this app can know: a scanned memo can drop a digit, and a typed job sheet can carry
      // a typo - on the 13-DM-26-0500 job both happened, and it was the typed sheet that
      // was wrong. So the differing pairs are shown and the choice is left to whoever has
      // the paper, rather than one source being assumed reliable.
      const sameCount = wanted.length === (inc.packages || []).length;
      const differences = sameCount
        ? (inc.packages || [])
            .map((p, i) => ({ from: p.code, to: wanted[i] }))
            .filter((d) => norm(d.from) !== norm(d.to))
        : [];
      const replaceable = missing.length > 0 && sameCount && differences.length > 0;
      return {
        codes: wanted.map((c) => (resolved.get(c) || {}).code).filter(Boolean),
        missing, replaceable, wanted, differences,
        elsewhere: [], text: byCode[codeKey].text || "", shared: [],
      };
    }
    const mark = sheetMarkForIncoming(inc);
    if (!mark || !(mark.numbers || []).length) return null;
    const sameLot = (code) => {
      if (!mark.lotCases) return true;
      const m = String(code).match(/\/(\d+)\s*$/);
      return !m || Number(m[1]) === mark.lotCases;
    };
    const done = new Set(inc.checkedInCodes || []);
    const available = (inc.packages || []).filter((p) => !done.has(p.code) && sameLot(p.code));
    const wanted = new Set(mark.numbers);
    const codes = available.filter((p) => wanted.has(codeLeadingNumber(p.code))).map((p) => p.code);
    const present = new Set(available.map((p) => codeLeadingNumber(p.code)));
    const missing = mark.numbers.filter((n) => !present.has(n));
    // A case the sheet asks for that this shipment doesn't hold may be on another shipment
    // for the same order - or may simply not have arrived. Point at the other shipment
    // without claiming it is the same box: the case numbers matching proves nothing on
    // their own, since each packing list numbers its own cases.
    const elsewhere = [];
    if (missing.length) {
      const myOrders = new Set(incomingOrders(inc));
      for (const other of matchedIncomings) {
        if (other.id === inc.id) continue;
        // Only a shipment that is plausibly the same lot is worth pointing at. Matching on
        // the case number alone listed every shipment on the site that happened to have a
        // case 1, 2 or 3 - which is nearly all of them, and none of them relevant.
        const sameOrder = incomingOrders(other).some((o) => myOrders.has(o));
        const sameLift = !!inc.unitCode && lotTokenMatches(inc.unitCode, other.unitCode);
        if (!sameOrder && !sameLift) continue;
        const otherDone = new Set(other.checkedInCodes || []);
        const found = (other.packages || [])
          .filter((p) => !otherDone.has(p.code) && sameLot(p.code) && missing.includes(codeLeadingNumber(p.code)))
          .map((p) => p.code);
        if (found.length) elsewhere.push({ inc: other, codes: found, sameOrder });
      }
    }
    const shared = shipmentsSharingMark(inc);
    return { codes, missing, elsewhere, text: mark.text || "", shared: shared.length > 1 ? shared : [] };
  }
  const incomingAutoApplied = row.incomingAutoApplied || {};
  const matchedIncomingKey = matchedIncomings.map((i) => i.id).join("|");
  useEffect(() => {
    if (!canMatchIncoming) return;
    const pending = {};
    for (const inc of matchedIncomings) {
      if (incomingAutoApplied[inc.id]) continue;
      if ((selectedByIncoming[inc.id] || []).length) continue;
      const sel = sheetSelectionForIncoming(inc);
      if (!sel || !sel.codes.length) continue;
      // Ambiguous: several shipments answer to this marking, so ticking either would be a
      // guess at which consignment the sheet is about.
      if (sel.shared.length) continue;
      pending[inc.id] = sel.codes;
    }
    const ids = Object.keys(pending);
    if (!ids.length) return;
    onChange({
      ...row,
      selectedByIncoming: { ...selectedByIncoming, ...pending },
      incomingAutoApplied: { ...incomingAutoApplied, ...Object.fromEntries(ids.map((id) => [id, true])) },
    });
    // Once per shipment, so a case unticked by hand is not ticked again on the next render.
  }, [matchedIncomingKey, canMatchIncoming]);
  // A Devan/CFS sheet names its lots and the cases in each, so where the packing list was
  // never filed - the older Gage Street jobs, whose account officer left before they were
  // entered - the sheet itself is enough to build one from. The cases come off the C/S NO.
  // marking and the lot's stated weight and volume are spread evenly across them, which is
  // all the sheet claims: it gives a figure per lot, never per case.
  // A delivery drawing on an arrival that is sitting in the same batch, waiting to be
  // processed. Arrivals are created before deliveries when Process runs, so that delivery
  // will find its entry - but the warning said flatly that no delivery would be recorded,
  // which reads as an instruction to go away and upload the Devan and CFS files first.
  function arrivalComingInThisBatch(jobNo) {
    const wanted = String(jobNo || "").trim();
    if (!wanted) return "";
    const hit = (siblingRows || []).find((r) => r !== row
      && ARRIVING_TYPES.includes(r.docType)
      && jobNosMatch(r.jobNumber, wanted));
    return hit ? legacySourceName(hit) : "";
  }
  const sheetLots = (row.refBlocks || []).flatMap((b) => b.lots || [])
    .map((lot) => {
      // The markings the sheet gave, as they were read. This used to re-parse the joined
      // text as a numeric case spec, which only succeeds where a lot happens to number its
      // cases 1, 2, 3 - so a Mitsubishi sheet listing "01C01, 04C10, 05C2105" produced no
      // lots at all, and the offer to build a packing list from the sheet never appeared on
      // any of them. With no packing list on file for the site, that left nothing to check
      // the cases into and every case had to be keyed in by hand.
      const codes = (lot.caseCodes || []).length
        ? lot.caseCodes.slice()
        : ((lot.caseNumbers || []).length
          ? lot.caseNumbers.map(String)
          : parseCaseSpec(lot.caseText || "").codes);
      return { lotRef: lot.lotRef || "", altRef: lot.altRef || "", unitCode: lot.unitCode || "", codes, kg: lot.kg || "", cbm: lot.cbm || "" };
    })
    .filter((lot) => lot.codes.length);
  function createPackingListFromSheet() {
    if (!onAddIncoming || !sheetLots.length) return;
    onAddIncoming(sheetLots.map((lot) => {
      const n = lot.codes.length;
      const per = (total, i) => {
        const value = Number(total) || 0;
        if (!value) return "";
        const each = Math.round((value / n) * 100) / 100;
        // The last case carries the rounding so the cases still add to the stated total.
        return String(i === n - 1 ? Math.round((value - each * (n - 1)) * 100) / 100 : each);
      };
      return {
        client: row.client,
        project: row.projectEn || row.projectZh || "",
        constructionSite: row.projectZh || "",
        jobRef: row.jobRef || "",
        shkNumber: row.shkNumber || "",
        unitCode: lot.unitCode,
        packages: lot.codes.map((code, i) => ({
          code, orderNo: lot.lotRef, description: "ELEVATOR PARTS",
          weightKg: per(lot.kg, i), cbm: per(lot.cbm, i),
        })),
        notes: t.legacyPackingListFromSheetNote(legacySourceName(row) || row.jobNumber || ""),
      };
    }));
  }
  const declaredIncomingGroups = groupDeclaredLots(row.declaredTotalsList, matchedIncomings);
  const declaredIncomingDist = distributeDeclaredAcrossLots(
    row.declaredTotalsList, matchedIncomings, incomingListedFor, row.declaredByGroup || {}
  );
  function getDeclaredFor(inc) {
    const found = declaredIncomingDist[inc.id];
    if (found && found.split) return { pkgs: found.pkgs, kg: found.kg, cbm: found.cbm };
    if (declaredByIncoming[inc.id] !== undefined) return declaredByIncoming[inc.id];
    return found ? { pkgs: found.pkgs, kg: found.kg, cbm: found.cbm } : { pkgs: "", kg: "", cbm: "" };
  }
  function setDeclaredFor(inc, patch) {
    const cur = getDeclaredFor(inc);
    onChange({ ...row, declaredByIncoming: { ...declaredByIncoming, [inc.id]: { ...cur, ...patch } } });
  }

  // Delivery: match against real inventory items (itemized, with cases still remaining)
  // instead of asking for a flat package count - prefer a typed "Refers to Arrival Job
  // No." if given, otherwise match by client + site the same way as Devan/CFS.
  // One sheet can close out several arrivals at once, so each "Ref Job no." block is
  // resolved on its own rather than the page being matched once as a whole. A block whose
  // job number finds nothing falls back to the lot codes it names (60766021/L52), which is
  // how an arrival filed under a different job number is still found; and a block that
  // resolves to nothing at all is reported rather than silently vanishing, which is what
  // used to happen - the sheet looked matched because its *other* block had matched.
  const deliveryRefBlocks = (() => {
    const fromSheet = (row.refBlocks || []).filter((b) => b.refJobNumber);
    const typed = String(row.referJobNumber || "").split(/[,;]/).map((s) => s.trim()).filter(Boolean);
    // A number typed in by hand that the sheet never mentioned still gets looked up.
    const extra = typed
      .filter((r) => !fromSheet.some((b) => jobNosMatch(b.refJobNumber, r)))
      .map((r) => ({ refJobNumber: r, lots: [] }));
    return [...fromSheet.filter((b) => typed.some((r) => jobNosMatch(b.refJobNumber, r))), ...extra];
  })();
  // What this sheet can pick from. A Delivery takes cases that are at the depot; a Return
  // brings back cases that are out at site, so it has to offer the delivered ones instead.
  // Offering it the depot's stock is why it found none of its cases and reported them all
  // as missing - they were missing precisely because they had been delivered, which is the
  // whole reason a Return exists.
  function selectablePackages(it) {
    if (row.docType !== "Return") return deliverablePackages(it);
    const out = new Set(deliveredCodes(it).map((c) => String(c).trim()));
    return (it.packages || []).filter((p) => out.has(String(p.code).trim()));
  }
  const deliveryMatch = (() => {
    if (!matchesInventory || !row.client) return { list: [], unmatched: [] };
    // A Delivery can only take what is still at the depot. A Return is putting cases back,
    // so the entries it needs are the ones that have been delivered from - filtering to what
    // is still deliverable would hide exactly the entries it is looking for.
    const base = (items || []).filter((it) => it.client === row.client
      && (it.packages || []).length > 0
      && (row.docType === "Return" || deliverablePackages(it).length > 0));
    const siteOk = (it) => !!(row.projectEn || row.projectZh)
      && sitesLooselyMatch(row.projectEn, row.projectZh, it.project, it.constructionSite);
    if (!deliveryRefBlocks.length) return { list: base.filter(siteOk), unmatched: [] };
    const chosen = new Map();
    const blockByItem = new Map();
    const unmatched = [];
    const take = (list, b) => list.forEach((it) => {
      chosen.set(it.id, it);
      // First block to claim an entry owns it. One sheet can name the same lot twice under
      // two different arrivals, and the cases each block lists belong to its own arrival.
      if (!blockByItem.has(it.id)) blockByItem.set(it.id, b);
    });
    for (const b of deliveryRefBlocks) {
      const byJob = base.filter((it) => jobNosMatch(it.jobNumber, b.refJobNumber));
      if (byJob.length) { take(byJob, b); continue; }
      const lotCodes = (b.lots || []).flatMap((l) => [l.unitCode, l.lotRef]).filter(Boolean);
      const byLot = lotCodes.length
        ? base.filter((it) => !chosen.has(it.id) && siteOk(it) && lotCodes.some((c) => lotTokenMatches(it.unitCode, c)))
        : [];
      if (byLot.length) { take(byLot, b); continue; }
      unmatched.push(b.refJobNumber);
    }
    return { list: [...chosen.values()], unmatched, blockByItem };
  })();
  const matchedItems = deliveryMatch.list;
  const unmatchedRefs = deliveryMatch.unmatched;
  const selectedByItem = row.selectedByItem || {};
  // The sheet already says which cases are leaving, so tick them instead of making the
  // user hunt for them among everything still at the depot. A case the sheet names but the
  // depot does not hold is reported rather than quietly dropped: this file reads "#3-7/34"
  // for L32-01 while case 3 has already gone, so four are ticked and case 3 is flagged -
  // which is exactly the "4 PKGS" the sheet declares for that block.
  // Mitsubishi cases are markings, not numbers - "01C3101-4-1" - so they are matched whole
  // against the codes the entry holds rather than by a leading number.
  // The lot this entry's own referral block names. Two blocks on one sheet can name the
  // same lot under two different arrivals - the 2607208 delivery closes L0MO-029239.002
  // twice, once against Devan 2606087 and once against 2607151 - and the maps keyed by lot
  // name merge the two lists into one and hand it to both entries. The block that matched
  // the entry is the only thing that says which of the two lists is this entry's, so it is
  // asked first and the maps are the fallback for a sheet that names no block at all.
  function blockLotFor(it) {
    const b = deliveryMatch.blockByItem && deliveryMatch.blockByItem.get(it.id);
    const lots = (b && b.lots) || [];
    if (!lots.length) return null;
    const byUnit = lots.find((l) => [l.unitCode, l.lotRef, l.altRef]
      .filter(Boolean).some((c) => lotTokenMatches(it.unitCode, c)));
    if (byUnit) return byUnit;
    // One arrival covering a single lot needs no matching; several, and only an exact lot
    // name will do, so the maps take over rather than a guess being made here.
    return lots.length === 1 ? lots[0] : null;
  }
  function sheetCodesFor(it) {
    const own = blockLotFor(it);
    if (own && (own.caseCodes || []).length) return { codes: own.caseCodes, text: own.caseText || "" };
    const byLot = row.caseCodesByLot || {};
    const keys = Object.keys(byLot);
    const hit = keys.find((k) => lotTokenMatches(it.unitCode, k))
      || keys.find((k) => jobNosMatch(it.jobNumber, k))
      || keys.find((k) => (it.packages || []).some((p) => lotTokenMatches(p.orderNo, k)));
    return hit ? byLot[hit] : null;
  }
  function sheetCasesFor(it) {
    const own = blockLotFor(it);
    if (own && (own.caseNumbers || []).length) {
      return { numbers: own.caseNumbers, text: own.caseText || "", lotCases: own.lotCases || null };
    }
    const byLot = row.caseMarksByLot || {};
    const lotKey = Object.keys(byLot).find((k) => lotTokenMatches(it.unitCode, k));
    if (lotKey) return byLot[lotKey];
    const byRef = row.caseMarksByRef || {};
    const refKey = Object.keys(byRef).find((k) => jobNosMatch(it.jobNumber, k));
    return refKey ? byRef[refKey] : null;
  }
  function sheetSelectionFor(it) {
    const coded = sheetCodesFor(it);
    if (coded && (coded.codes || []).length) {
      const deliverable = selectablePackages(it);
      // fall through to the shared shape below
      // Matched with the brackets off, so a sheet's "(10-1/10B-1)" finds the case the
      // packing list holds as "10-1/10B-1".
      const codes = coded.codes
        .map((c) => (resolveCaseCode(c, deliverable) || {}).code)
        .filter(Boolean);
      const missing = coded.codes.filter((c) => !resolveCaseCode(c, deliverable));
      return { codes, missing, text: coded.text || "" };
    }
    const mark = sheetCasesFor(it);
    if (!mark || !(mark.numbers || []).length) return null;
    // A case number only means something together with its lot size: "#4,6/20" is case 4
    // of a twenty-case lot, and case 4/19 sitting in the same entry is a different box on
    // a different lift. Where the sheet states the lot size, cases numbered against a
    // different one are ignored rather than picked up by their leading number.
    const sameLot = (code) => {
      if (!mark.lotCases) return true;
      const m = String(code).match(/\/(\d+)\s*$/);
      return !m || Number(m[1]) === mark.lotCases;
    };
    const deliverable = selectablePackages(it).filter((p) => sameLot(p.code));
    const wanted = new Set(mark.numbers);
    const codes = deliverable.filter((p) => wanted.has(codeLeadingNumber(p.code))).map((p) => p.code);
    const present = new Set(deliverable.map((p) => codeLeadingNumber(p.code)));
    return { codes, missing: mark.numbers.filter((n) => !present.has(n)), text: mark.text || "" };
  }
  // The Schindler booking for 60766021/60766022 mislabels its LIFT NO. column on four
  // rows, so cases 4/20 and 6/20 were checked in under L52 and 18/19 and 19/19 under L53.
  // The importer now reads that correctly, but entries created before it did are still
  // wrong, and rebuilding them from the packing list forward is a lot to ask.
  //
  // The case number itself gives the proof: every case of a nineteen-case lot is numbered
  // "n/19". So an entry whose cases are overwhelmingly out of 19, sitting beside one whose
  // cases are overwhelmingly out of 20, can claim any stray "n/19" from its neighbour -
  // and it works in both directions, whether or not this particular delivery happens to
  // name the case. An entry with no clear lot size, or a neighbour with the same one, is
  // left alone.
  function dominantLotSize(it) {
    const counts = new Map();
    for (const p of it.packages || []) {
      const m = String(p.code).match(/\/(\d+)\s*$/);
      if (!m) continue;
      const n = Number(m[1]);
      counts.set(n, (counts.get(n) || 0) + 1);
    }
    if (!counts.size) return null;
    const ranked = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    if (ranked.length > 1 && ranked[0][1] === ranked[1][1]) return null;
    return ranked[0][0];
  }
  // A case the sheet asks for that isn't on this entry is usually somewhere - checked in
  // under the twin lift, or under the lot the job sheet and the packing list name
  // differently. Saying "not at the depot" when it is at the depot, one row down, is the
  // difference between a puzzle and a correction someone can make in a few seconds. Only
  // this client's entries are searched, and markings are compared with brackets off so a
  // sheet's "(10-1/10B-1)" finds the case the packing list holds as "10-1/10B-1".
  function locateMissingCases(missing, exceptId) {
    const out = [];
    for (const code of missing || []) {
      for (const other of items || []) {
        if (other.id === exceptId || other.cancelled) continue;
        if (row.client && other.client && other.client !== row.client) continue;
        if (!(other.packages || []).some((p) => sameCaseCode(p.code, code))) continue;
        out.push({ code, label: `${other.id}${other.unitCode ? ` \u00b7 ${other.unitCode}` : ""}` });
        break;
      }
    }
    return out;
  }
  function misfiledCasesFor(it) {
    const mine = dominantLotSize(it);
    if (!mine) return [];
    const here = new Set((it.packages || []).map((p) => p.code));
    const found = [];
    for (const other of matchedItems) {
      if (other.id === it.id) continue;
      const theirs = dominantLotSize(other);
      if (!theirs || theirs === mine) continue;
      for (const p of other.packages || []) {
        const m = String(p.code).match(/\/(\d+)\s*$/);
        if (!m || Number(m[1]) !== mine || here.has(p.code)) continue;
        found.push({ code: p.code, from: other });
      }
    }
    return found.sort((a, b) => codeLeadingNumber(a.code) - codeLeadingNumber(b.code));
  }
  function repairMisfiled(target, moves) {
    if (!onLegacyEnrich || !moves.length) return;
    const bySource = new Map();
    for (const m of moves) {
      if (!bySource.has(m.from.id)) bySource.set(m.from.id, { item: m.from, codes: [] });
      bySource.get(m.from.id).codes.push(m.code);
    }
    const entries = [];
    const moved = [];
    const nextSelected = { ...selectedByItem };
    for (const { item, codes } of bySource.values()) {
      const set = new Set(codes);
      moved.push(...(item.packages || []).filter((p) => set.has(p.code)));
      entries.push({
        itemId: item.id,
        patch: {
          packages: (item.packages || []).filter((p) => !set.has(p.code)),
          // The case leaves its old entry's arrival batches too, or that entry would go on
          // believing it took delivery of a box it never held.
          arrivals: (item.arrivals || []).map((a) => ({ ...a, codes: (a.codes || []).filter((c) => !set.has(c)) })),
        },
      });
      nextSelected[item.id] = (nextSelected[item.id] || []).filter((c) => !set.has(c));
    }
    const denom = (c) => { const m = String(c).match(/\/(\d+)\s*$/); return m ? Number(m[1]) : 0; };
    const patch = {
      packages: [...(target.packages || []), ...moved]
        .sort((a, b) => denom(a.code) - denom(b.code) || codeLeadingNumber(a.code) - codeLeadingNumber(b.code)),
    };
    // An entry tracked by arrival batches would treat an unlisted case as not yet landed,
    // so the moved cases join its earliest batch and stay deliverable.
    if (usesArrivalBatches(target)) {
      const codes = moved.map((p) => p.code);
      patch.arrivals = (target.arrivals || []).map((a, i) =>
        i === 0 ? { ...a, codes: [...new Set([...(a.codes || []), ...codes])] } : a);
    }
    entries.push({ itemId: target.id, patch });
    onLegacyEnrich(entries);
    nextSelected[target.id] = [...new Set([...(nextSelected[target.id] || []), ...moved.map((p) => p.code)])];
    onChange({ ...row, selectedByItem: nextSelected });
  }
  const caseAutoApplied = row.caseAutoApplied || {};
  const matchedItemKey = matchedItems.map((it) => it.id).join("|");
  useEffect(() => {
    if (row.docType !== "Delivery") return;
    const pending = {};
    for (const it of matchedItems) {
      if (caseAutoApplied[it.id]) continue;
      if ((selectedByItem[it.id] || []).length) continue;
      const sel = sheetSelectionFor(it);
      if (!sel || !sel.codes.length) continue;
      pending[it.id] = sel.codes;
    }
    const ids = Object.keys(pending);
    if (!ids.length) return;
    onChange({
      ...row,
      selectedByItem: { ...selectedByItem, ...pending },
      caseAutoApplied: { ...caseAutoApplied, ...Object.fromEntries(ids.map((id) => [id, true])) },
    });
    // Runs once per matched item: the applied flag survives further edits, so a case the
    // user then unticks by hand is not silently ticked again on the next render.
  }, [matchedItemKey, row.docType]);
  // A Delivery sheet declares its own totals just as a Devan does - the L7/L8/L9 sheet
  // closes with one line covering all twelve packages - and that figure is the accurate
  // one, so it drives the delivery instead of the packing-list sum of the cases ticked.
  const declaredByItem = row.declaredByItem || {};
  const declaredByGroup = row.declaredByGroup || {};
  const itemListedFor = (it) => sumSelectedPackages(it.packages, selectedByItem[it.id] || []);
  const declaredItemGroups = groupDeclaredLots(row.declaredTotalsList, matchedItems);
  const declaredItemDist = distributeDeclaredAcrossLots(
    row.declaredTotalsList, matchedItems, itemListedFor, declaredByGroup
  );
  // A total the sheet states once for several lots is edited once, as the sheet wrote it.
  // Each lot's share is derived from it, so correcting the sheet figure re-splits them
  // all rather than leaving the user to reconcile three boxes by hand.
  function getDeclaredGroupTotal(group) {
    const key = declaredGroupKey(group);
    if (declaredByGroup[key] !== undefined) return declaredByGroup[key];
    return { pkgs: group.line.pkgs || "", kg: group.line.kg || "", cbm: group.line.cbm || "" };
  }
  function setDeclaredGroupTotal(group, patch) {
    const key = declaredGroupKey(group);
    onChange({ ...row, declaredByGroup: { ...declaredByGroup, [key]: { ...getDeclaredGroupTotal(group), ...patch } } });
  }
  function getDeclaredForItem(it) {
    const found = declaredItemDist[it.id];
    // Lots inside a shared total are driven by that total, not by their own box.
    if (found && found.split) return { pkgs: found.pkgs, kg: found.kg, cbm: found.cbm };
    if (declaredByItem[it.id] !== undefined) return declaredByItem[it.id];
    return found ? { pkgs: found.pkgs, kg: found.kg, cbm: found.cbm } : { pkgs: "", kg: "", cbm: "" };
  }
  function setDeclaredForItem(it, patch) {
    const cur = getDeclaredForItem(it);
    onChange({ ...row, declaredByItem: { ...declaredByItem, [it.id]: { ...cur, ...patch } } });
  }
  // Oversize on the way out. A delivery sheet lists its own OVERSIZE CASES section, so
  // the cases and CBM leaving on this job are recorded against the delivery rather than
  // inferred from whatever was flagged when the lot arrived.
  const oversizeByDeliveryItem = row.oversizeByDeliveryItem || {};
  function getDeliveryOversizeFor(it) {
    if (oversizeByDeliveryItem[it.id] !== undefined) return normaliseOversize(oversizeByDeliveryItem[it.id]);
    const detected = (row.oversizeByLot || {})[it.unitCode];
    return detected && detected.length
      ? { checked: true, cases: detected.map((c) => ({ code: c.code || "", cbm: c.cbm || "" })) }
      : { checked: false, cases: [{ code: "", cbm: "" }] };
  }
  function setDeliveryOversizeFor(it, patch) {
    const cur = normaliseOversize(getDeliveryOversizeFor(it));
    onChange({ ...row, oversizeByDeliveryItem: { ...oversizeByDeliveryItem, [it.id]: { ...cur, ...patch } } });
  }
  // Picking cases one tap at a time is slow on a lot of thirty-four. Three shortcuts sit
  // alongside the typed "1,3-5,7" box: tapping one case and then shift-tapping another
  // takes everything between them, Clear empties the lot, and tapping a description takes
  // every case of that kind at once - which is how these deliveries are usually described
  // in the first place ("the landing doors and the mech").
  const [lastTappedByItem, setLastTappedByItem] = useState({});
  function toggleItemCode(itemId, code, index, extend, remainingPkgs) {
    const cur = selectedByItem[itemId] || [];
    const from = lastTappedByItem[itemId];
    let next;
    if (extend && from != null && remainingPkgs && from !== index) {
      const span = remainingPkgs.slice(Math.min(from, index), Math.max(from, index) + 1).map((p) => p.code);
      next = [...new Set([...cur, ...span])];
    } else {
      next = toggleOccurrence(remainingPkgs || [], index, cur);
    }
    setLastTappedByItem((prev) => ({ ...prev, [itemId]: index }));
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: next } });
  }
  function selectAllItem(itemId, remainingPkgs) {
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: remainingPkgs.map((p) => p.code) } });
  }
  function clearItem(itemId) {
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: [] } });
  }
  function toggleItemDescription(itemId, description, remainingPkgs) {
    const codes = remainingPkgs.filter((p) => (p.description || "") === description).map((p) => p.code);
    const cur = selectedByItem[itemId] || [];
    const allOn = codes.every((c) => cur.includes(c));
    const next = allOn ? cur.filter((c) => !codes.includes(c)) : [...new Set([...cur, ...codes])];
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: next } });
  }
  const [typeSelectText, setTypeSelectText] = useState({});
  function applyTypeSelectItem(itemId, remainingPkgs) {
    const nums = parseRangeInput(typeSelectText[itemId]);
    if (nums.size === 0) return;
    const matchingCodes = remainingPkgs.filter((p) => nums.has(codeLeadingNumber(p.code))).map((p) => p.code);
    const cur = selectedByItem[itemId] || [];
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: [...new Set([...cur, ...matchingCodes])] } });
    setTypeSelectText((prev) => ({ ...prev, [itemId]: "" }));
  }
  function applyTypeSelectIncoming(incId, remainingPkgs) {
    const nums = parseRangeInput(typeSelectText[incId]);
    if (nums.size === 0) return;
    const matchingCodes = remainingPkgs.filter((p) => nums.has(codeLeadingNumber(p.code))).map((p) => p.code);
    const cur = selectedByIncoming[incId] || [];
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: [...new Set([...cur, ...matchingCodes])] } });
    setTypeSelectText((prev) => ({ ...prev, [incId]: "" }));
  }
  return (
    <div className="rounded p-3 flex flex-col gap-3" style={{ border: `1px solid ${colors.line}`, background: colors.surface }}>
      <div className="flex items-start justify-between gap-2">
        <div className="text-sm font-semibold" style={{ color: colors.ink, wordBreak: "break-all" }}>{legacySourceName(row)}</div>
        <button type="button" className="text-xs font-semibold whitespace-nowrap" style={{ color: colors.red }} onClick={onRemove}>{t.deleteBtn}</button>
      </div>
      {row.autoDetected && (
        <div className="px-2 py-1 rounded text-xs w-fit" style={{ background: colors.amberSoft, color: colors.amberText }}>
          {t.legacyAutoDetectedTag}
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
        <Field label={t.legacyDocType} colors={colors}>
          <select className={inputClass} style={inputStyle} value={row.docType} onChange={set("docType")}>
            {LEGACY_DOC_TYPES.map((tp) => <option key={tp} value={tp}>{tp}</option>)}
          </select>
        </Field>
        <Field label={t.clientLabel} colors={colors}>
          <select className={inputClass} style={{ ...inputStyle, borderColor: !row.client ? colors.red : inputStyle.borderColor }} value={row.client} onChange={set("client")}>
            <option value="">{t.legacyClientUnresolved}</option>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.legacyProjectSiteEn} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, borderColor: !row.projectEn && !row.projectZh ? colors.red : inputStyle.borderColor }} value={row.projectEn} onChange={set("projectEn")} />
        </Field>
        <Field label={t.legacyProjectSiteZh} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, borderColor: !row.projectEn && !row.projectZh ? colors.red : inputStyle.borderColor }} value={row.projectZh} onChange={set("projectZh")} />
        </Field>
        <div className="col-span-1 sm:col-span-2 md:col-span-4">
          {/* The Directory entry behind this sheet's site, found by matching the name the
              sheet uses. Editing it here saves the walk to the Directory tab mid-upload. */}
          <InlineSiteEditor
            site={(directory || []).find((d) => sitesLooselyMatch(row.projectEn, row.projectZh, d.siteEn, d.siteZh))}
            setDirectory={setDirectory} employees={employees} colors={colors} t={t}
          />
        </div>
        <Field label={t.fJobNumber} colors={colors}>
          <input className={inputClass} style={inputStyle} value={row.jobNumber} onChange={set("jobNumber")} />
        </Field>
        <Field label={t.colDate} colors={colors}>
          {/* Ringed while empty. Everything downstream is ordered by this date, so a blank
              one is not a small omission - it is a record that sorts to the end of the
              ledger and bills from nothing. */}
          <input type="date" className={inputClass}
            style={{ ...inputStyle, ...(String(row.date || "").trim() ? {} : { borderColor: colors.red, background: colors.redSoft }) }}
            value={row.date} onChange={set("date")} />
        </Field>
        {itemized && (
          <>
            {(matchedIncomings.length > 0 || matchedItems.length > 0) ? null : (
              <Field label={t.legacyUnitCode} colors={colors}>
                <input className={inputClass} style={inputStyle} value={row.unitCode} onChange={set("unitCode")} />
              </Field>
            )}
            {(matchedIncomings.length > 0 || matchedItems.length > 0) ? null : (
              <>
                <Field label={t.legacyPkgs} colors={colors}>
                  <input type="number" min="0" className={inputClass} style={inputStyle} value={row.packageCount} onChange={set("packageCount")} />
                </Field>
                <Field label={t.legacyWeightKg} colors={colors}>
                  <input type="number" min="0" className={inputClass} style={inputStyle} value={row.weightKg} onChange={set("weightKg")} />
                </Field>
                <Field label={t.legacyCbm} colors={colors}>
                  <input type="number" min="0" step="0.001" className={inputClass} style={inputStyle} value={row.volumeCbm} onChange={set("volumeCbm")} />
                </Field>
              </>
            )}
            {(row.docType === "Devan" || row.docType === "CFS") && (
              <>
                <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.shkNumber} onChange={set("shkNumber")} />
                </Field>
                <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.jobRef} onChange={set("jobRef")} />
                </Field>
                <div className="col-span-2 md:col-span-3">
                  <Field label={t.fSsDoNo} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={row.ssDoNo} onChange={set("ssDoNo")} placeholder={'ex ss."SHIP" V.___; CONTAINERS NO. ___'} />
                  </Field>
                </div>
              </>
            )}
            {matchesInventory && (
              <>
                <Field label={t.legacyReferJobNoLabel} hint={t.legacyReferJobNoHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.referJobNumber} onChange={set("referJobNumber")} />
                </Field>
                <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.shkNumber} onChange={set("shkNumber")} />
                </Field>
              </>
            )}
          </>
        )}
      </div>
      {onProcessAll && (matchedIncomings.length > 0 || matchedItems.length > 0) && (
        <button
          type="button"
          className="px-4 py-2 rounded text-sm font-semibold w-fit"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: processDisabled ? 0.5 : 1 }}
          disabled={processDisabled}
          onClick={onProcessAll}
        >
          {processing ? t.legacyProcessingMsg : t.saveBtn}
        </button>
      )}
      {canMatchIncoming && matchedIncomings.length === 0 && matchedItems.length === 0 && sheetLots.length > 0 && onAddIncoming && (
        <div className="rounded p-3" style={{ background: colors.amberSoft, border: `1px solid ${colors.amber}` }}>
          <div className="text-sm font-semibold mb-1" style={{ color: colors.amberText }}>{t.legacyNoPackingListTitle}</div>
          {/* A flat "nothing on file" is a claim the app cannot always support: a packing
              list may well be here and simply not have matched, and building a second one
              from the sheet would then double the site's cases. What was looked for, and
              what is actually waiting, is reported instead - so a client or site written
              differently on one side shows itself. */}
          <div className="text-xs mb-2" style={{ color: colors.amberText }}>{t.legacyPackingListLookedFor(row.client, row.projectEn || row.projectZh)}</div>
          {nearMissIncomings.length > 0 && (
            <div className="text-xs mb-2 rounded px-2 py-1.5" style={{ background: colors.redSoft, color: colors.red }}>
              {t.legacyPackingListNearMiss(nearMissIncomings.length, nearMissIncomings.slice(0, 4)
                .map((inc) => `${inc.id} \u00b7 ${inc.client || "?"} \u00b7 ${inc.project || inc.constructionSite || "?"}`).join("; "))}
            </div>
          )}
          <div className="text-xs mb-2" style={{ color: colors.amberText }}>{t.legacyNoPackingListDesc}</div>
          {sheetLots.map((lot, li) => (
            <div key={li} className="text-xs flex flex-wrap gap-3" style={{ color: colors.ink }}>
              <span style={{ fontFamily: FONT_MONO, minWidth: 190 }}>{lot.lotRef}{lot.unitCode ? ` / ${lot.unitCode}` : ""}</span>
              <span style={{ minWidth: 190 }}>{lot.codes.join(", ")}</span>
              <span style={{ fontFamily: FONT_MONO }}>{lot.kg ? `${lot.kg} kg` : "\u2014"} · {lot.cbm ? `${lot.cbm} cbm` : "\u2014"}</span>
            </div>
          ))}
          <button
            type="button"
            className="px-3 py-1.5 rounded text-xs font-semibold mt-2"
            style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
            onClick={createPackingListFromSheet}
          >
            {t.legacyCreatePackingListBtn(sheetLots.length)}
          </button>
        </div>
      )}
      {matchedIncomings.length > 0 && (
        <div className="rounded p-3" style={{ background: colors.greenSoft, border: `1px solid ${colors.green}` }}>
          <div className="text-sm font-semibold mb-2" style={{ color: colors.green }}>
            {t.legacyMatchedIncomingCount(matchedIncomings.length)}
          </div>
          <div className="mb-3 max-w-xs">
            <Field label={t.packingListApplyDepot} colors={colors}>
              <select className={inputClass} style={inputStyle} value={row.depot || DEPOTS[0]} onChange={set("depot")}>
                {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
              </select>
            </Field>
          </div>
          {declaredIncomingGroups.filter((g) => g.lots.length > 1).map((g) => (
            <SharedDeclaredTotal
              key={declaredGroupKey(g)}
              group={g}
              value={(row.declaredByGroup || {})[declaredGroupKey(g)] || { pkgs: g.line.pkgs || "", kg: g.line.kg || "", cbm: g.line.cbm || "" }}
              onPatch={(patch) => {
                const key = declaredGroupKey(g);
                const cur = (row.declaredByGroup || {})[key] || { pkgs: g.line.pkgs || "", kg: g.line.kg || "", cbm: g.line.cbm || "" };
                onChange({ ...row, declaredByGroup: { ...(row.declaredByGroup || {}), [key]: { ...cur, ...patch } } });
              }}
              colors={colors} t={t} inputClass={inputClass} inputStyle={inputStyleFor(colors)}
            />
          ))}
          <div className="flex flex-col gap-4">
            {matchedIncomings.map((inc) => {
              const done = new Set(inc.checkedInCodes || []);
              const remainingPkgs = (inc.packages || []).filter((p) => !done.has(p.code));
              const selectedCodes = selectedByIncoming[inc.id] || [];
              const sheetSel = sheetSelectionForIncoming(inc);
              const totals = sumSelectedPackages(inc.packages, selectedCodes);
              const oversize = getOversizeFor(inc);
              const declared = getDeclaredFor(inc);
              const variance = computeDeclaredVariance(totals, declared);
              // Two Incoming records can hold the same lot under different labels - the
              // Gage Street site has 60701670 filed once as "L5" and again as "L5 1st
              // batch", and the cases this file is looking for sit in only one of them.
              // Where a sibling shares an order number, say so and say what it holds, so
              // the right one can be picked rather than guessed at.
              const myOrders = new Set((inc.packages || []).map((p) => String(p.orderNo || "").trim()).filter(Boolean));
              const sameOrderSiblings = myOrders.size
                ? matchedIncomings.filter((other) => other.id !== inc.id
                    && (other.packages || []).some((p) => myOrders.has(String(p.orderNo || "").trim())))
                : [];
              return (
                <div key={inc.id} style={{ borderTop: `1px solid ${colors.green}`, paddingTop: 10 }}>
                  <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
                    <div className="text-xs font-semibold" style={{ color: colors.green, fontFamily: FONT_DISPLAY }}>
                      {t.legacyMatchedIncoming(inc.id)} {"\u00b7"} {incomingLabel(inc)}
                    </div>
                    <div className="flex items-center gap-2">
                      <input
                        className={inputClass}
                        style={{ ...inputStyleFor(colors), width: 140, fontSize: 12, padding: "4px 8px" }}
                        placeholder={t.legacyTypeSelectPlaceholder}
                        value={typeSelectText[inc.id] || ""}
                        onChange={(e) => setTypeSelectText((prev) => ({ ...prev, [inc.id]: e.target.value }))}
                        onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); applyTypeSelectIncoming(inc.id, remainingPkgs); } }}
                      />
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => applyTypeSelectIncoming(inc.id, remainingPkgs)}>{t.legacyTypeSelectBtn}</button>
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => selectAllIncoming(inc.id, remainingPkgs)}>{t.selectAllBtn}</button>
                    </div>
                  </div>
                  <div className="mb-2">
                    <OversizeCasesEditor
                      value={oversize}
                      onToggle={(checked) => setOversizeFor(inc, { checked })}
                      onCases={(cases) => setOversizeFor(inc, { cases })}
                      colors={colors} t={t} inputClass={inputClass} inputStyle={inputStyleFor(colors)}
                    />
                  </div>
                  {selectedCodes.length > 0 && (
                    <div className="mb-2">
                      <div className="text-xs font-semibold mb-1.5" style={{ color: colors.ink }}>
                        {t.legacySelectedTotals(totals.count, Math.round(totals.weight * 10) / 10, Math.round(totals.cbm * 1000) / 1000)}
                      </div>
                      {declaredIncomingDist[inc.id] && declaredIncomingDist[inc.id].split ? (
                        <div className="text-xs mb-1" style={{ color: colors.inkFaint }}>
                          {t.legacyDeclaredShareNote(
                            Math.round((Number(declared.kg) || 0) * 10) / 10,
                            Math.round((Number(declared.cbm) || 0) * 1000) / 1000
                          )}
                        </div>
                      ) : (
                        <div className="flex flex-wrap items-center gap-2 mb-1">
                          <div className="text-xs font-semibold" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                            {t.legacyDeclaredLabel}
                          </div>
                          {[
                            { key: "pkgs", ph: t.jsPkgs, w: 64 },
                            { key: "kg", ph: t.jsKgs, w: 90 },
                            { key: "cbm", ph: t.jsCbm, w: 82 },
                          ].map((f) => (
                            <input
                              key={f.key}
                              type="number" min="0" step="0.001"
                              className={inputClass}
                              style={{ ...inputStyleFor(colors), width: f.w, fontSize: 12, padding: "4px 8px" }}
                              placeholder={f.ph}
                              value={declared[f.key] || ""}
                              onChange={(e) => setDeclaredFor(inc, { [f.key]: e.target.value })}
                            />
                          ))}
                        </div>
                      )}
                      {variance && variance.any && (
                        <div className="text-xs" style={{ color: variance.pkgs && variance.pkgs.delta !== 0 ? colors.amberText : colors.inkFaint }}>
                          {variance.pkgs && variance.pkgs.delta !== 0 ? t.legacyDeclaredPkgsGap(variance.pkgs.declared, variance.pkgs.listed) : ""}
                          {variance.kg && Math.abs(variance.kg.pct) >= 0.05
                            ? `${variance.pkgs && variance.pkgs.delta !== 0 ? " " : ""}${t.legacyDeclaredKgGap(Math.round(Math.abs(variance.kg.delta) * 10) / 10, variance.kg.delta > 0, Math.abs(variance.kg.pct).toFixed(1))}`
                            : ""}
                        </div>
                      )}
                      <div className="text-xs" style={{ color: colors.inkFaint }}>{t.legacyDeclaredHint}</div>
                    </div>
                  )}
                  {sheetSel && (
                    <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>
                      {sheetSel.shared.length > 0
                        ? t.legacySheetCasesAmbiguous(sheetSel.text, sheetSel.shared.map((x) => x.id).join(", "))
                        : t.legacySheetCasesNote(sheetSel.text, sheetSel.codes.length)}
                      {sheetSel.missing.length > 0 && (
                        <span style={{ color: colors.red }}> {"\u00b7"} {t.legacyIncomingCasesMissing(sheetSel.missing.join(", "))}</span>
                      )}
                      {/* Nothing matched at all. Almost always the two documents write the
                          same case a different way - a job sheet's "01B1101" against a
                          packing list's "B11 01" - and without seeing both side by side
                          there is nothing to tell you that, only a screen full of cases to
                          tick by hand. */}
                      {sheetSel.codes.length === 0 && sheetSel.missing.length > 0 && (
                        <div style={{ color: colors.red }}>
                          {t.legacyCasesFormatClash(
                            sheetSel.missing.slice(0, 3).join(", "),
                            remainingPkgs.slice(0, 3).map((p) => p.code).join(", ")
                          )}
                        </div>
                      )}
                      {sheetSel.replaceable && onReplaceIncomingCases && (
                        <div style={{ color: colors.amberText }}>
                          <div>{t.legacyReplaceCasesHint(sheetSel.differences.length)}</div>
                          <div style={{ fontFamily: FONT_MONO }}>
                            {sheetSel.differences.map((d) => `${d.from} \u2192 ${d.to}`).join("  \u00b7  ")}
                          </div>
                          <button
                            type="button"
                            className="font-semibold underline"
                            onClick={() => onReplaceIncomingCases(inc.id, sheetSel.wanted)}
                          >
                            {t.legacyReplaceCasesBtn}
                          </button>
                        </div>
                      )}
                      {sheetSel.elsewhere.slice(0, 4).map((e) => (
                        <div key={e.inc.id} style={{ color: colors.amberText }}>
                          {t.legacyCasesFoundIn(e.codes.join(", "), e.inc.id, incomingLabel(e.inc))}
                        </div>
                      ))}
                      {sheetSel.elsewhere.length > 4 && (
                        <div style={{ color: colors.amberText }}>{t.legacyCasesFoundInMore(sheetSel.elsewhere.length - 4)}</div>
                      )}
                    </div>
                  )}
                  {sameOrderSiblings.length > 0 && (
                    <div className="px-2 py-1.5 rounded text-xs mb-2" style={{ background: colors.amberSoft, color: colors.amberText }}>
                      {t.legacySameOrderWarn([...myOrders].join(", "))}
                      {sameOrderSiblings.map((other) => {
                        const otherDone = new Set(other.checkedInCodes || []);
                        const otherLeft = (other.packages || []).filter((p) => !otherDone.has(p.code));
                        const mine = new Set(remainingPkgs.map((p) => p.code));
                        const onlyThere = otherLeft.filter((p) => !mine.has(p.code)).map((p) => p.code);
                        return (
                          <div key={other.id}>
                            {t.legacySameOrderSibling(other.id, incomingLabel(other), otherLeft.length)}
                            {onlyThere.length > 0 && ` \u00b7 ${t.legacySameOrderOnlyThere(onlyThere.join(", "))}`}
                          </div>
                        );
                      })}
                    </div>
                  )}
                  <div className="flex flex-col gap-3">
                    {groupPackagesByOrder(remainingPkgs).map((grp) => (
                      <div key={grp.orderNo || "_"}>
                        {grp.orderNo && (
                          <div className="text-xs font-semibold mb-1.5" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                            {grp.orderNo}{inc.unitCode ? ` / ${inc.unitCode}` : ""}
                          </div>
                        )}
                        <div className="flex flex-wrap gap-2">
                          {grp.packages.map((p) => {
                            // Chips are keyed and highlighted by position, not by marking,
                            // so two cases sharing a marking can be picked apart.
                            const gi = remainingPkgs.findIndex((x) => x === p);
                            const on = occurrenceSelected(remainingPkgs, gi, selectedCodes);
                            return (
                            <button
                              key={`${p.code}-${gi}`}
                              type="button"
                              onClick={() => setSelectedForIncoming(inc.id, toggleOccurrence(remainingPkgs, gi, selectedCodes))}
                              className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                              style={{
                                border: `1px solid ${on ? colors.amber : colors.line}`,
                                background: on ? colors.amberSoft : colors.surface,
                                color: on ? colors.amberText : colors.ink,
                              }}
                              title={p.description}
                            >
                              {p.code}{p.description ? ` \u2014 ${p.description}` : ""}
                            </button>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
      {matchedItems.length > 0 && (
        <div className="rounded p-3" style={{ background: colors.greenSoft, border: `1px solid ${colors.green}` }}>
          <div className="text-sm font-semibold mb-2" style={{ color: colors.green }}>
            {row.docType === "Return" ? t.legacyMatchedItemsReturn(matchedItems.length) : t.legacyMatchedItemsCount(matchedItems.length)}
          </div>
          {(() => {
            const grand = matchedItems.reduce((acc, it) => {
              const codes = selectedByItem[it.id] || [];
              const s = effectiveDeliveryTotals(selectedItemTotals(it, codes), getDeclaredForItem(it));
              return { count: acc.count + s.count, weight: acc.weight + s.weight, cbm: acc.cbm + s.cbm };
            }, { count: 0, weight: 0, cbm: 0 });
            return grand.count > 0 ? (
              <div className="text-xs mb-3 font-semibold px-2 py-1.5 rounded" style={{ color: colors.ink, background: colors.surface, width: "fit-content" }}>
                {t.legacySelectedTotalsGrand(grand.count, Math.round(grand.weight * 10) / 10, Math.round(grand.cbm * 1000) / 1000)}
              </div>
            ) : null;
          })()}
          {declaredItemGroups.filter((g) => g.lots.length > 1).map((g) => (
            <SharedDeclaredTotal
              key={declaredGroupKey(g)}
              group={g}
              value={getDeclaredGroupTotal(g)}
              onPatch={(patch) => setDeclaredGroupTotal(g, patch)}
              colors={colors} t={t} inputClass={inputClass} inputStyle={inputStyleFor(colors)}
            />
          ))}
          {matchedItems.length > 0 && (
            <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>{t.legacyRangeTapHint}</div>
          )}
          <div className="flex flex-col gap-4">
            {matchedItems.map((it) => {
              const remainingPkgs = selectablePackages(it);
              const sheetSel = sheetSelectionFor(it);
              const selectedCodes = selectedByItem[it.id] || [];
              const derivedTotals = selectedItemTotals(it, selectedCodes);
              const itemDeclared = getDeclaredForItem(it);
              const totals = effectiveDeliveryTotals(derivedTotals, itemDeclared);
              const itemVariance = computeDeclaredVariance(derivedTotals, itemDeclared);
              const dOversize = getDeliveryOversizeFor(it);
              return (
                <div key={it.id} style={{ borderTop: `1px solid ${colors.green}`, paddingTop: 10 }}>
                  <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
                    <div className="text-xs font-semibold" style={{ color: colors.green, fontFamily: FONT_DISPLAY }}>
                      {row.docType === "Return" ? t.legacyMatchedItemReturn(it.id) : t.legacyMatchedItem(it.id)}{it.unitCode ? ` \u00b7 ${it.unitCode}` : ""}{it.jobNumber ? ` \u00b7 ${t.fJobNumber}: ${it.jobNumber}` : ""}
                    </div>
                    <div className="flex items-center gap-2">
                      <input
                        className={inputClass}
                        style={{ ...inputStyleFor(colors), width: 140, fontSize: 12, padding: "4px 8px" }}
                        placeholder={t.legacyTypeSelectPlaceholder}
                        value={typeSelectText[it.id] || ""}
                        onChange={(e) => setTypeSelectText((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); applyTypeSelectItem(it.id, remainingPkgs); } }}
                      />
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => applyTypeSelectItem(it.id, remainingPkgs)}>{t.legacyTypeSelectBtn}</button>
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => selectAllItem(it.id, remainingPkgs)}>{t.selectAllBtn}</button>
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => clearItem(it.id)}>{t.clearBtn}</button>
                    </div>
                  </div>
                  {sheetSel && (
                    <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>
                      {t.legacySheetCasesNote(sheetSel.text, sheetSel.codes.length)}
                      {sheetSel.missing.length > 0 && (
                        <span style={{ color: colors.red }}> {"\u00b7"} {t.legacySheetCasesMissing(sheetSel.missing.join(", "))}</span>
                      )}
                      {(() => {
                        const found = locateMissingCases(sheetSel.missing, it.id);
                        if (!found.length) return null;
                        return (
                          <span style={{ color: colors.amberText }}> {"\u00b7"} {t.legacySheetCasesElsewhere(
                            found.slice(0, 6).map((f) => `${f.code} \u2192 ${f.label}`).join(", "),
                            found.length > 6 ? found.length - 6 : 0
                          )}</span>
                        );
                      })()}
                    </div>
                  )}
                  {(() => {
                    const misfiled = misfiledCasesFor(it);
                    if (!misfiled.length) return null;
                    const from = [...new Set(misfiled.map((m) => m.from.unitCode || m.from.id))].join(", ");
                    return (
                      <div className="px-2 py-1.5 rounded text-xs mb-2 flex flex-wrap items-center gap-2"
                        style={{ background: colors.amberSoft, color: colors.amberText }}>
                        <span>{t.legacyMisfiledCases(misfiled.map((m) => m.code).join(", "), from)}</span>
                        <button type="button" className="font-semibold underline"
                          onClick={() => repairMisfiled(it, misfiled)}>
                          {t.legacyMisfiledFixBtn}
                        </button>
                      </div>
                    );
                  })()}
                  {selectedCodes.length > 0 && (
                    <div className="mb-2">
                      <div className="text-xs font-semibold mb-1.5" style={{ color: colors.ink }}>
                        {t.legacySelectedTotals(totals.count, Math.round(totals.weight * 10) / 10, Math.round(totals.cbm * 1000) / 1000)}
                      </div>
                      {declaredItemDist[it.id] && declaredItemDist[it.id].split ? (
                        <div className="text-xs mb-1" style={{ color: colors.inkFaint }}>
                          {t.legacyDeclaredShareNote(
                            Math.round((Number(itemDeclared.kg) || 0) * 10) / 10,
                            Math.round((Number(itemDeclared.cbm) || 0) * 1000) / 1000
                          )}
                        </div>
                      ) : (
                        <div className="flex flex-wrap items-center gap-2 mb-1">
                          <div className="text-xs font-semibold" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                            {t.legacyDeliveryDeclaredLabel}
                          </div>
                          {[
                            { key: "kg", ph: t.jsKgs, w: 90 },
                            { key: "cbm", ph: t.jsCbm, w: 82 },
                          ].map((f) => (
                            <input
                              key={f.key}
                              type="number" min="0" step="0.001"
                              className={inputClass}
                              style={{ ...inputStyleFor(colors), width: f.w, fontSize: 12, padding: "4px 8px" }}
                              placeholder={f.ph}
                              value={itemDeclared[f.key] || ""}
                              onChange={(e) => setDeclaredForItem(it, { [f.key]: e.target.value })}
                            />
                          ))}
                        </div>
                      )}
                      {itemVariance && itemVariance.kg && Math.abs(itemVariance.kg.pct) >= 0.05 && (
                        <div className="text-xs" style={{ color: colors.inkFaint }}>
                          {t.legacyDeliveryDeclaredGap(
                            Math.round(itemVariance.kg.declared * 10) / 10,
                            Math.round(itemVariance.kg.listed * 10) / 10,
                            Math.abs(itemVariance.kg.pct).toFixed(1),
                            itemVariance.kg.delta > 0
                          )}
                        </div>
                      )}
                      <div className="mt-1.5">
                        <OversizeCasesEditor
                          value={dOversize}
                          onToggle={(checked) => setDeliveryOversizeFor(it, { checked })}
                          onCases={(cases) => setDeliveryOversizeFor(it, { cases })}
                          colors={colors} t={t} inputClass={inputClass} inputStyle={inputStyleFor(colors)}
                        />
                      </div>
                    </div>
                  )}
                  {(() => {
                    const kinds = [];
                    for (const p of remainingPkgs) {
                      const d = p.description || "";
                      if (!d) continue;
                      const hit = kinds.find((k) => k.description === d);
                      if (hit) hit.count += 1; else kinds.push({ description: d, count: 1 });
                    }
                    if (kinds.length < 2) return null;
                    return (
                      <div className="flex flex-wrap gap-1.5 mb-2">
                        {kinds.map((k) => {
                          const codes = remainingPkgs.filter((p) => (p.description || "") === k.description).map((p) => p.code);
                          const allOn = codes.every((c) => selectedCodes.includes(c));
                          return (
                            <button
                              key={k.description}
                              type="button"
                              onClick={() => toggleItemDescription(it.id, k.description, remainingPkgs)}
                              className="px-2 py-1 rounded text-xs"
                              style={{
                                border: `1px dashed ${allOn ? colors.amber : colors.line}`,
                                background: allOn ? colors.amberSoft : "transparent",
                                color: allOn ? colors.amberText : colors.inkFaint,
                              }}
                            >
                              {k.description} · {k.count}
                            </button>
                          );
                        })}
                      </div>
                    );
                  })()}
                  <div className="flex flex-wrap gap-2">
                    {remainingPkgs.map((p, idx) => (
                      <button
                        key={`${p.code}-${idx}`}
                        type="button"
                        onClick={(e) => toggleItemCode(it.id, p.code, idx, e.shiftKey, remainingPkgs)}
                        className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                        style={{
                          border: `1px solid ${occurrenceSelected(remainingPkgs, idx, selectedCodes) ? colors.amber : colors.line}`,
                          background: occurrenceSelected(remainingPkgs, idx, selectedCodes) ? colors.amberSoft : colors.surface,
                          color: occurrenceSelected(remainingPkgs, idx, selectedCodes) ? colors.amberText : colors.ink,
                        }}
                        title={p.description}
                      >
                        {p.code}{p.description ? ` \u2014 ${p.description}` : ""}
                      </button>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
      {!row.projectEn && !row.projectZh && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {t.legacySiteRequiredMsg}
        </div>
      )}
      {(row.docType === "Devan" || row.docType === "CFS") && (
        <div className="text-xs" style={{ color: colors.inkFaint }}>{t.legacyArrivalStaysOpenHint}</div>
      )}
      {row.scanError && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {t.legacyScanFailed(row.scanError)}
        </div>
      )}
      {row.scannedFromPdf && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
          {t.legacyScannedFromPdf}
        </div>
      )}
      {(row.caseCountMismatches || []).length > 0 && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {(row.caseCountMismatches || []).map((m, i) => (
            <div key={i}>{t.legacyCaseCountMismatch(m.lot, m.stated, m.listed)}</div>
          ))}
        </div>
      )}
      {row.hasPastedContentImage && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
          {t.legacyCasesArePicture}
        </div>
      )}
      {legacyRowMissing(row) && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {t.legacyRowMissingHint(legacyRowMissing(row).map((k) => t.legacyFieldNames[k]).join(", "))}
        </div>
      )}
      {matchesInventory && matchedItems.length === 0 && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {row.referJobNumber
            ? (arrivalComingInThisBatch(row.referJobNumber)
              ? t.legacyArrivalLaterInBatch(row.referJobNumber, arrivalComingInThisBatch(row.referJobNumber))
              : t.legacyNoArrivalFoundHint(row.referJobNumber))
            : t.legacyNoReferralHint}
        </div>
      )}
      {matchesInventory && matchedItems.length > 0 && unmatchedRefs.length > 0 && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {t.legacySomeArrivalsUnmatched(unmatchedRefs.join(", "))}
        </div>
      )}
    </div>
  );
}

// The manual packing list lives with Incoming because that is what it produces: a shipment
// waiting to be checked in, exactly as an uploaded packing list would leave one. It is for
// the older jobs whose packing list was never filed - the Gage Street work whose account
// officer left before the records were entered - so the Devan/CFS sheet has something to
// check its cases into.
function ManualPackingListEntry({ onClose, onAddIncoming, existingItems, directory, setDirectory, employees, colors, t }) {
  const [showOlderSites, setShowOlderSites] = useState(false);
  const [manualForm, setManualForm] = useState({
    client: CLIENTS[0], project: "", constructionSite: "", orderedBy: "", jobRef: "", shkNumber: "", unitCode: "", directoryId: "", saveToDirectory: false, packages: [],
  });
  // A job sheet names several lots under one heading - the 2512079 CFS covers four - so
  // lots finished here are parked in this list while the next is entered, and all of them
  // go to Incoming together.
  const [manualLots, setManualLots] = useState([]);
  const [manualLotDraft, setManualLotDraft] = useState({ orderNo: "", caseSpec: "", description: "ELEVATOR PARTS", kg: "", cbm: "" });
  const inputStyle = inputStyleFor(colors);
  // Directory sites first: those are the ones that carry a Chinese name, job ref and
  // ordered-by to fill in behind them. Site names only ever seen on existing entries come
  // after, so a job that was never added to the directory can still be typed.
  const siteSuggestions = useMemo(() => {
    const fromDirectory = (directory || []).map((s) => s.siteEn).filter(Boolean);
    const fromItems = (existingItems || []).map((i) => i.project).filter(Boolean);
    return [...new Set([...fromDirectory, ...fromItems.sort()])];
  }, [directory, existingItems]);
  // Picking a name from the list is the same as picking it from the directory select
  // above: matching it back to its directory entry is what fills the Chinese name, job ref
  // and ordered-by, and what stops the form offering to save a site that already exists.
  // Matching ignores case, punctuation and a leading "SITE AT", since a name typed off a
  // job sheet rarely matches the directory character for character.
  function findDirectorySite(name) {
    const key = normalizeSiteForMatch(name);
    if (!key) return null;
    const hit = (d) => normalizeSiteForMatch(d.siteEn) === key || normalizeSiteForMatch(d.siteZh) === key;
    return (directory || []).find((d) => d.client === manualForm.client && hit(d))
      || (directory || []).find(hit) || null;
  }
  function applyProjectName(value) {
    const site = findDirectorySite(value);
    setManualForm((f) => (site ? {
      ...f,
      project: site.siteEn || value,
      directoryId: site.id,
      client: site.client || f.client,
      constructionSite: site.siteZh || f.constructionSite,
      jobRef: site.jobRef || f.jobRef,
      orderedBy: site.orderedBy || f.orderedBy,
      saveToDirectory: false,
    } : { ...f, project: value, directoryId: "" }));
  }
  // Builds a lot's cases from the C/S NO. marking as written on the sheet, spreading the
  // lot's stated weight and volume evenly across them. A job sheet gives a figure per lot,
  // not per case, and an even split is the only honest reading of that - the per-case
  // figures stay editable afterwards for anyone who knows better.
  function buildLotCases() {
    const { codes } = parseCaseSpec(manualLotDraft.caseSpec);
    if (!codes.length) return [];
    const n = codes.length;
    const kg = Number(manualLotDraft.kg) || 0;
    const cbm = Number(manualLotDraft.cbm) || 0;
    const per = (total, i) => {
      if (!total) return "";
      // The last case carries the rounding so the cases still add up to the stated total.
      const each = Math.round((total / n) * 100) / 100;
      return String(i === n - 1 ? Math.round((total - each * (n - 1)) * 100) / 100 : each);
    };
    return codes.map((code, i) => ({
      code, orderNo: manualLotDraft.orderNo || "",
      description: manualLotDraft.description || "",
      weightKg: per(kg, i), cbm: per(cbm, i),
    }));
  }
  // One action per lot: the cases are made from the C/S NO. marking, or taken from the
  // itemized list below for a lot whose codes are not plain numbers, and the lot is parked
  // in the list. A CFS sheet naming four lots is four presses and one submit, rather than
  // four separate packing lists.
  function addLot() {
    const built = buildLotCases();
    const packages = built.length ? built : (manualForm.packages || []);
    if (!packages.length) return;
    setManualLots((prev) => [...prev, {
      unitCode: manualForm.unitCode || "",
      orderNo: manualLotDraft.orderNo || "",
      packages,
    }]);
    setManualForm((f) => ({ ...f, unitCode: "", packages: [] }));
    setManualLotDraft({ orderNo: "", caseSpec: "", description: "ELEVATOR PARTS", kg: "", cbm: "" });
  }
  const manualPendingCount = manualLots.length + ((manualForm.packages || []).length ? 1 : 0);
  function addManualToIncoming() {
    let effectiveDirectoryId = manualForm.directoryId || "";
    if (manualForm.saveToDirectory && !manualForm.directoryId && manualForm.project) {
      const newSite = {
        id: `SITE${Date.now()}`,
        siteEn: manualForm.project,
        siteZh: manualForm.constructionSite && manualForm.constructionSite !== manualForm.project ? manualForm.constructionSite : "",
        client: manualForm.client,
        jobRef: manualForm.jobRef || "",
        orderedBy: manualForm.orderedBy || "",
        accountOfficer: "",
      };
      setDirectory((d) => [...d, newSite]);
      effectiveDirectoryId = newSite.id;
    }
    const lots = [
      ...manualLots,
      ...((manualForm.packages || []).length
        ? [{ unitCode: manualForm.unitCode || "", orderNo: manualLotDraft.orderNo || "", packages: manualForm.packages }]
        : []),
    ];
    onAddIncoming(lots.map((lot) => ({
      client: manualForm.client,
      project: manualForm.project,
      constructionSite: manualForm.constructionSite || "",
      jobRef: manualForm.jobRef || "",
      orderedBy: manualForm.orderedBy || "",
      shkNumber: manualForm.shkNumber || "",
      directoryId: effectiveDirectoryId,
      unitCode: lot.unitCode,
      packages: lot.packages,
      notes: t.legacyManualEntryNote,
    })));
    setManualForm({ client: CLIENTS[0], project: "", constructionSite: "", orderedBy: "", jobRef: "", shkNumber: "", unitCode: "", directoryId: "", saveToDirectory: false, packages: [] });
    setManualLots([]);
    setManualLotDraft({ orderNo: "", caseSpec: "", description: "ELEVATOR PARTS", kg: "", cbm: "" });
    // Collapse once the shipments exist, so it is clear the entry went through.
    if (onClose) onClose();
  }
  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center justify-end">
        <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={onClose}>{t.cancelBtn}</button>
      </div>
        <div className="flex flex-col gap-4">
        <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
          <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.tabManualPackingList}</h3>
          <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>{t.manualPackingListDesc}</p>
          <div className="mb-4">
            <Field label={t.packingListFillFromDirectory} colors={colors}>
              <select
                className={inputClass}
                style={{ ...inputStyle, minWidth: 320 }}
                value={manualForm.directoryId || ""}
                onChange={(e) => {
                  const site = (directory || []).find((d) => d.id === e.target.value);
                  if (!site) { setManualForm((f) => ({ ...f, directoryId: "" })); return; }
                  setManualForm((f) => ({
                    ...f,
                    directoryId: site.id,
                    client: site.client || f.client,
                    project: site.siteEn || f.project,
                    constructionSite: site.siteZh || site.siteEn || f.constructionSite,
                    jobRef: site.jobRef || f.jobRef,
                    orderedBy: site.orderedBy || f.orderedBy,
                    saveToDirectory: false,
                  }));
                }}
              >
                <option value="">{t.selectFromDirectoryPlaceholder}</option>
                {visibleDirectory(directory, { client: manualForm.client, showOlder: showOlderSites, items: existingItems }).map((site) => (
                  <option key={site.id} value={site.id}>{site.siteEn} \u2014 {site.client}</option>
                ))}
              </select>
            </Field>
            <label className="flex items-center gap-1.5 text-xs mt-1.5" style={{ color: colors.inkFaint }}>
              <input type="checkbox" checked={showOlderSites} onChange={(e) => setShowOlderSites(e.target.checked)} />
              {t.showOlderJobs}
              {!showOlderSites && hiddenSiteCount(directory, { client: manualForm.client, items: existingItems }) > 0 && (
                <span>{t.showOlderJobsCount(hiddenSiteCount(directory, { client: manualForm.client, items: existingItems }))}</span>
              )}
            </label>
            <InlineSiteEditor
              site={(directory || []).find((d) => d.id === manualForm.directoryId)}
              setDirectory={setDirectory} employees={employees} colors={colors} t={t}
            />
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
            <Field label={t.packingListApplyClient} colors={colors}>
              <select className={inputClass} style={inputStyle} value={manualForm.client} onChange={(e) => setManualForm((f) => ({ ...f, client: e.target.value, directoryId: "" }))}>
                {CLIENTS.map((c) => <option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label={t.legacyProjectSiteEn} colors={colors}>
              <input list="manual-site-suggestions" className={inputClass} style={inputStyle} value={manualForm.project} onChange={(e) => applyProjectName(e.target.value)} />
              <datalist id="manual-site-suggestions">
                {siteSuggestions.map((s) => <option key={s} value={s} />)}
              </datalist>
            </Field>
            <Field label={t.legacyProjectSiteZh} colors={colors}>
              <input className={inputClass} style={inputStyle} value={manualForm.constructionSite} onChange={(e) => setManualForm((f) => ({ ...f, constructionSite: e.target.value }))} />
            </Field>
            <Field label={t.fOrderedBy} colors={colors}>
              <input className={inputClass} style={inputStyle} value={manualForm.orderedBy} onChange={(e) => setManualForm((f) => ({ ...f, orderedBy: e.target.value }))} />
            </Field>
            <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
              <input className={inputClass} style={inputStyle} value={manualForm.jobRef} onChange={(e) => setManualForm((f) => ({ ...f, jobRef: e.target.value }))} />
            </Field>
            <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
              <input className={inputClass} style={inputStyle} value={manualForm.shkNumber} onChange={(e) => setManualForm((f) => ({ ...f, shkNumber: e.target.value }))} />
            </Field>

          </div>
          {manualForm.directoryId ? (
            <div className="mt-3 text-xs" style={{ color: colors.green }}>{t.manualLinkedToDirectory}</div>
          ) : manualForm.project ? (
            <label className="flex items-center gap-2 mt-4 text-sm" style={{ color: colors.inkFaint }}>
              <input type="checkbox" checked={manualForm.saveToDirectory} onChange={(e) => setManualForm((f) => ({ ...f, saveToDirectory: e.target.checked }))} />
              {t.saveNewSiteToDirectory(manualForm.project)}
            </label>
          ) : null}
          <div className="rounded p-3 mt-4" style={{ border: `1px dashed ${colors.line}` }}>
            <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
              {t.manualLotBuilderLabel}
            </div>
            <div className="text-[11px] mb-2" style={{ color: colors.inkFaint }}>{t.manualLotBuilderHint}</div>
            <div className="flex flex-wrap items-end gap-2">
              <Field label={t.legacyUnitCode} colors={colors}>
                <input className={inputClass} style={{ ...inputStyle, width: 120 }} value={manualForm.unitCode}
                  onChange={(e) => setManualForm((f) => ({ ...f, unitCode: e.target.value }))} placeholder="P3" />
              </Field>
              <Field label={t.manualOrderNoLabel} colors={colors}>
                <input className={inputClass} style={{ ...inputStyle, width: 130 }} value={manualLotDraft.orderNo}
                  onChange={(e) => setManualLotDraft((d) => ({ ...d, orderNo: e.target.value }))} placeholder="60737177" />
              </Field>
              <Field label={t.manualCaseSpecLabel} colors={colors}>
                <input className={inputClass} style={{ ...inputStyle, width: 150 }} value={manualLotDraft.caseSpec}
                  onChange={(e) => setManualLotDraft((d) => ({ ...d, caseSpec: e.target.value }))} placeholder="1,2,3/3" />
              </Field>
              <Field label={t.fDescription} colors={colors}>
                <input className={inputClass} style={{ ...inputStyle, width: 180 }} value={manualLotDraft.description}
                  onChange={(e) => setManualLotDraft((d) => ({ ...d, description: e.target.value }))} />
              </Field>
              <Field label={t.jsKgs} colors={colors}>
                <input type="number" min="0" step="0.01" className={inputClass} style={{ ...inputStyle, width: 110 }} value={manualLotDraft.kg}
                  onChange={(e) => setManualLotDraft((d) => ({ ...d, kg: e.target.value }))} />
              </Field>
              <Field label={t.jsCbm} colors={colors}>
                <input type="number" min="0" step="0.001" className={inputClass} style={{ ...inputStyle, width: 100 }} value={manualLotDraft.cbm}
                  onChange={(e) => setManualLotDraft((d) => ({ ...d, cbm: e.target.value }))} />
              </Field>
              <button className="px-3 py-2 rounded text-xs font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
                onClick={addLot}>
                {t.manualAddLotBtn}
              </button>
            </div>
            {parseCaseSpec(manualLotDraft.caseSpec).codes.length > 0 && (
              <div className="text-xs mt-2" style={{ color: colors.green }}>
                {t.manualCaseSpecPreview(parseCaseSpec(manualLotDraft.caseSpec).codes.length, parseCaseSpec(manualLotDraft.caseSpec).codes.join(", "))}
              </div>
            )}
          </div>
          <PackagesEditor form={manualForm} setForm={setManualForm} colors={colors} t={t} />
          {manualLots.length > 0 && (
            <div className="rounded p-3 mt-3" style={{ background: colors.surfaceDim }}>
              <div className="text-xs font-semibold mb-1" style={{ color: colors.ink }}>{t.manualPendingLotsLabel(manualLots.length)}</div>
              {manualLots.map((lot, li) => (
                <div key={li} className="text-xs flex items-center gap-3" style={{ color: colors.ink }}>
                  <span style={{ fontFamily: FONT_MONO, minWidth: 150 }}>{[lot.orderNo, lot.unitCode].filter(Boolean).join(" / ") || t.manualUnnamedLot}</span>
                  <span className="flex-1">{lot.packages.map((p) => p.code).join(", ")}</span>
                  <button className="font-semibold" style={{ color: colors.red }}
                    onClick={() => setManualLots((prev) => prev.filter((_, i) => i !== li))}>{t.deleteBtn}</button>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="flex gap-2">
          <button
            className="px-4 py-2 rounded text-sm font-semibold w-fit"
            style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: (!manualForm.client || !(manualForm.project || manualForm.constructionSite) || manualPendingCount === 0) ? 0.5 : 1 }}
            disabled={!manualForm.client || !(manualForm.project || manualForm.constructionSite) || manualPendingCount === 0}
            onClick={addManualToIncoming}
          >
            {t.packingListAddToIncomingBtn(manualPendingCount)}
          </button>
        </div>
      </div>

    </div>
  );
}
function IncomingPanel({ incoming, setIncoming, items, directory, setDirectory, employees, onCheckIn, onAddIncoming, colors, t, lang }) {
  const [search, setSearch] = useState("");
  const [filterClient, setFilterClient] = useState("All");
  const [showCompleted, setShowCompleted] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [selectedByShipment, setSelectedByShipment] = useState({});
  const [formByShipment, setFormByShipment] = useState({});
  const [manualOpen, setManualOpen] = useState(false);
  // Clearing a site's shipments one card at a time is the same hundred clicks the backlog
  // had. Selection follows the filters, so "select all" is always the cards on screen and
  // never one a filter is hiding.
  const [picked, setPicked] = useState({});
  const inputStyle = inputStyleFor(colors);

  function remainingPkgs(inc) {
    const done = new Set(inc.checkedInCodes || []);
    return (inc.packages || []).filter((p) => !done.has(p.code));
  }
  function isComplete(inc) {
    return remainingPkgs(inc).length === 0;
  }

  const clientOptions = useMemo(() => [...new Set(incoming.map((i) => i.client).filter(Boolean))].sort(), [incoming]);
  const pickedIncoming = incoming.filter((i) => picked[i.id]);
  const pickedCases = pickedIncoming.reduce((n, i) => n + (i.packages || []).length, 0);
  // A shipment already checked into an inventory entry is the one worth warning about:
  // deleting it leaves that entry with no packing list behind it.
  const pickedLinked = pickedIncoming.filter((i) => i.linkedItemId).length;
  const filtered = incoming.filter((inc) => {
    if (filterClient !== "All" && inc.client !== filterClient) return false;
    if (!showCompleted && isComplete(inc)) return false;
    if (!search.trim()) return true;
    const q = search.toLowerCase();
    // Searchable by the things written on the paperwork - the SHK reference and the
    // order/commission number on the cases - as well as by site, lift and case code.
    return inc.client?.toLowerCase().includes(q) || inc.project?.toLowerCase().includes(q) ||
      inc.constructionSite?.toLowerCase().includes(q) || inc.unitCode?.toLowerCase().includes(q) ||
      inc.shkNumber?.toLowerCase().includes(q) || inc.jobRef?.toLowerCase().includes(q) ||
      (inc.packages || []).some((p) => (p.code || "").toLowerCase().includes(q)
        || String(p.orderNo || "").toLowerCase().includes(q));
  });

  function getSel(id) { return selectedByShipment[id] || []; }
  function toggleCode(id, code) {
    setSelectedByShipment((prev) => {
      const cur = prev[id] || [];
      return { ...prev, [id]: cur.includes(code) ? cur.filter((c) => c !== code) : [...cur, code] };
    });
  }
  function selectAll(inc) {
    setSelectedByShipment((prev) => ({ ...prev, [inc.id]: remainingPkgs(inc).map((p) => p.code) }));
  }
  function getForm(id) {
    return formByShipment[id] || { type: "Devan", depot: DEPOTS[0], jobNumber: "", date: todayStr(), ssDoNo: "", declaredPkgs: "", declaredKg: "", declaredCbm: "" };
  }
  function setForm(id, patch) {
    setFormByShipment((prev) => ({ ...prev, [id]: { ...getForm(id), ...patch } }));
  }

  function submitCheckIn(inc) {
    const codes = getSel(inc.id);
    const form = getForm(inc.id);
    if (codes.length === 0 || !form.jobNumber || !form.date) return;
    const declared = (form.declaredPkgs || form.declaredKg || form.declaredCbm)
      ? { pkgs: form.declaredPkgs, kg: form.declaredKg, cbm: form.declaredCbm }
      : null;
    onCheckIn({
      incomingId: inc.id, codes, type: form.type, depot: form.depot,
      jobNumber: form.jobNumber, date: form.date, ssDoNo: form.ssDoNo,
      declared, declaredSource: form.type,
    });
    setSelectedByShipment((prev) => ({ ...prev, [inc.id]: [] }));
    setFormByShipment((prev) => ({ ...prev, [inc.id]: { ...getForm(inc.id), jobNumber: "", declaredPkgs: "", declaredKg: "", declaredCbm: "" } }));
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <div className="flex items-start justify-between gap-3 flex-wrap">
          <div>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.incomingTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.incomingDesc}</p>
          </div>
          {!manualOpen && (
            <button
              className="px-4 py-2 rounded text-sm font-semibold whitespace-nowrap"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => setManualOpen(true)}
            >
              {t.manualPackingListOpenBtn}
            </button>
          )}
        </div>
        {manualOpen && (
          <div className="mb-4">
            <ManualPackingListEntry
              onClose={() => setManualOpen(false)}
              onAddIncoming={onAddIncoming} existingItems={items} employees={employees}
              directory={directory} setDirectory={setDirectory} employees={employees} colors={colors} t={t}
            />
          </div>
        )}
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.searchLabel} colors={colors}>
            <input className={inputClass} style={{ ...inputStyle, minWidth: 220 }} value={search} onChange={(e) => setSearch(e.target.value)} />
          </Field>
          <Field label={t.clientLabel} colors={colors}>
            <select className={inputClass} style={inputStyle} value={filterClient} onChange={(e) => setFilterClient(e.target.value)}>
              <option>All</option>
              {clientOptions.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <label className="flex items-center gap-2 text-sm pb-2" style={{ color: colors.inkFaint }}>
            <input type="checkbox" checked={showCompleted} onChange={(e) => setShowCompleted(e.target.checked)} />
            {t.incomingShowCompleted}
          </label>
        </div>
      </div>

      <div className="flex flex-col gap-3">
        {filtered.length > 0 && (
          <div className="flex flex-wrap items-center gap-3 px-1">
            <label className="flex items-center gap-2 text-xs cursor-pointer" style={{ color: colors.inkFaint }}>
              <input type="checkbox"
                checked={filtered.every((i) => picked[i.id])}
                onChange={(e) => {
                  const next = { ...picked };
                  filtered.forEach((i) => { if (e.target.checked) next[i.id] = true; else delete next[i.id]; });
                  setPicked(next);
                }} />
              {t.incomingSelectAll(filtered.length)}
            </label>
            {pickedIncoming.length > 0 && (
              <>
                <span className="text-xs font-semibold" style={{ color: colors.ink }}>{t.incomingSelectedCount(pickedIncoming.length, pickedCases)}</span>
                <button className="text-xs font-semibold" style={{ color: colors.red }}
                  onClick={() => {
                    if (!window.confirm(t.incomingBulkDeleteConfirm(pickedIncoming.length, pickedCases, pickedLinked))) return;
                    const ids = new Set(pickedIncoming.map((i) => i.id));
                    setIncoming((prev) => prev.filter((i) => !ids.has(i.id)));
                    setPicked({});
                  }}>
                  {t.incomingBulkDeleteBtn}
                </button>
                <button className="text-xs" style={{ color: colors.inkFaint }} onClick={() => setPicked({})}>{t.legacyClearSelection}</button>
              </>
            )}
          </div>
        )}
        {filtered.length === 0 && (
          <div className="px-3 py-6 text-center text-sm rounded-lg" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.inkFaint }}>
            {t.incomingNoneMsg}
          </div>
        )}
        {filtered.map((inc) => {
          const remaining = remainingPkgs(inc);
          const complete = isComplete(inc);
          const isOpen = expandedId === inc.id;
          const sel = getSel(inc.id);
          const form = getForm(inc.id);
          return (
            <div key={inc.id} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}`, background: colors.surface }}>
              <div
                className="px-4 py-3 flex flex-wrap items-center justify-between gap-2 cursor-pointer"
                style={{ background: colors.surfaceDim }}
                onClick={() => setExpandedId(isOpen ? null : inc.id)}
              >
                <div className="flex items-start gap-3">
                  <input type="checkbox" className="mt-1" checked={!!picked[inc.id]}
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => setPicked((p) => ({ ...p, [inc.id]: e.target.checked }))} />
                  <div>
                  <div className="text-sm font-bold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
                    {inc.client} · {inc.project || inc.constructionSite}{inc.unitCode ? ` \u00b7 ${inc.unitCode}` : ""}
                  </div>
                  <div className="text-xs" style={{ color: colors.inkFaint }}>
                    {[
                      inc.shkNumber,
                      // The order/commission numbers the cases carry. A shipment is looked
                      // up by these as often as by its lift code, and several lots can sit
                      // under one lift name, so they belong on the card.
                      [...new Set((inc.packages || []).map((p) => String(p.orderNo || "").trim()).filter(Boolean))].slice(0, 3).join(", "),
                      t.incomingCaseCount((inc.packages || []).length),
                      inc.linkedItemId ? t.incomingLinkedTo(inc.linkedItemId) : "",
                    ].filter(Boolean).join(" \u00b7 ")}
                  </div>
                </div>
                </div>
                {complete ? (
                  <Badge tone="green" colors={colors}>{t.incomingFullyCheckedIn}</Badge>
                ) : (
                  <Badge tone="amber" colors={colors}>{t.incomingRemainingBadge(remaining.length, (inc.packages || []).length)}</Badge>
                )}
                <button
                  type="button"
                  title={t.incomingDeleteBtn}
                  aria-label={t.incomingDeleteBtn}
                  className="w-6 h-6 rounded-full inline-flex items-center justify-center font-bold"
                  style={{ background: colors.redSoft, color: colors.red, lineHeight: 1 }}
                  onClick={(e) => { e.stopPropagation(); if (window.confirm(t.incomingDeleteConfirm)) setIncoming((prev) => prev.filter((i) => i.id !== inc.id)); }}
                >
                  &minus;
                </button>
              </div>
              {isOpen && (
                <div className="p-4 flex flex-col gap-4">
                  <div className="max-w-xs">
                    <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                      <input
                        className={inputClass}
                        style={inputStyleFor(colors)}
                        value={inc.shkNumber || ""}
                        onChange={(e) => setIncoming((prev) => prev.map((i) => (i.id === inc.id ? { ...i, shkNumber: e.target.value } : i)))}
                      />
                    </Field>
                  </div>
                  {remaining.length > 0 ? (
                    <>
                      <div>
                        <div className="flex items-center justify-between mb-2">
                          <div className="text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{t.incomingSelectCasesLabel}</div>
                          <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => selectAll(inc)}>{t.selectAllBtn}</button>
                        </div>
                        <div className="flex flex-col gap-3">
                          {groupPackagesByOrder(remaining).map((grp) => (
                            <div key={grp.orderNo || "_"}>
                              {grp.orderNo && (
                                <div className="text-xs font-semibold mb-1.5" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                                  {grp.orderNo}{inc.unitCode ? ` / ${inc.unitCode}` : ""}
                                </div>
                              )}
                              <div className="flex flex-wrap gap-2">
                                {grp.packages.map((p) => (
                                  <button
                                    key={p.code}
                                    type="button"
                                    onClick={() => toggleCode(inc.id, p.code)}
                                    className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                                    style={{
                                      border: `1px solid ${sel.includes(p.code) ? colors.amber : colors.line}`,
                                      background: sel.includes(p.code) ? colors.amberSoft : colors.surface,
                                      color: sel.includes(p.code) ? colors.amberText : colors.ink,
                                    }}
                                    title={p.description}
                                  >
                                    {p.code}{p.description ? ` \u2014 ${p.description}` : ""}
                                  </button>
                                ))}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
                        <Field label={t.fArrivingType} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={form.type} onChange={(e) => setForm(inc.id, { type: e.target.value })}>
                            {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
                          </select>
                        </Field>
                        <Field label={t.packingListApplyDepot} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={form.depot} onChange={(e) => setForm(inc.id, { depot: e.target.value })}>
                            {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                          </select>
                        </Field>
                        <Field label={t.colDate} colors={colors}>
                          <input type="date" className={inputClass} style={inputStyle} value={form.date} onChange={(e) => setForm(inc.id, { date: e.target.value })} />
                        </Field>
                        <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
                          <div className="flex gap-2">
                            <input className={inputClass + " flex-1"} style={inputStyle} value={form.jobNumber} onChange={(e) => setForm(inc.id, { jobNumber: e.target.value })} />
                            <button
                              type="button"
                              className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
                              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
                              onClick={() => setForm(inc.id, { jobNumber: nextJobNumber(items) })}
                            >
                              {t.generateJobNoBtn}
                            </button>
                          </div>
                        </Field>
                        {form.type === "Devan" && (
                          <div className="col-span-1 sm:col-span-2 md:col-span-4">
                            <Field label={t.fSsDoNo} hint={t.fSsDoNoHint} colors={colors}>
                              <input className={inputClass} style={inputStyle} value={form.ssDoNo} onChange={(e) => setForm(inc.id, { ssDoNo: e.target.value })} placeholder={'ex ss."SHIP" V.___; CONTAINERS NO. ___'} />
                            </Field>
                          </div>
                        )}
                        <div className="col-span-1 sm:col-span-2 md:col-span-4">
                          {(() => {
                            const selTotals = sumSelectedPackages(inc.packages, getSel(inc.id));
                            const declaredNow = { pkgs: form.declaredPkgs, kg: form.declaredKg, cbm: form.declaredCbm };
                            const v = computeDeclaredVariance(selTotals, declaredNow);
                            return (
                              <Field label={t.incomingDeclaredLabel} hint={t.legacyDeclaredHint} colors={colors}>
                                <div className="flex flex-wrap items-center gap-2">
                                  {[
                                    { key: "declaredPkgs", ph: t.jsPkgs, w: 70 },
                                    { key: "declaredKg", ph: t.jsKgs, w: 96 },
                                    { key: "declaredCbm", ph: t.jsCbm, w: 88 },
                                  ].map((f) => (
                                    <input
                                      key={f.key}
                                      type="number" min="0" step="0.001"
                                      className={inputClass}
                                      style={{ ...inputStyle, width: f.w }}
                                      placeholder={f.ph}
                                      value={form[f.key] || ""}
                                      onChange={(e) => setForm(inc.id, { [f.key]: e.target.value })}
                                    />
                                  ))}
                                  {v && v.any && (
                                    <span className="text-xs" style={{ color: v.pkgs && v.pkgs.delta !== 0 ? colors.amberText : colors.inkFaint }}>
                                      {v.pkgs && v.pkgs.delta !== 0 ? t.legacyDeclaredPkgsGap(v.pkgs.declared, v.pkgs.listed) : ""}
                                      {v.kg && Math.abs(v.kg.pct) >= 0.05
                                        ? `${v.pkgs && v.pkgs.delta !== 0 ? " " : ""}${t.legacyDeclaredKgGap(Math.round(Math.abs(v.kg.delta) * 10) / 10, v.kg.delta > 0, Math.abs(v.kg.pct).toFixed(1))}`
                                        : ""}
                                    </span>
                                  )}
                                </div>
                              </Field>
                            );
                          })()}
                        </div>
                      </div>
                      <button
                        className="px-4 py-2 rounded text-sm font-semibold w-fit"
                        style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: (sel.length === 0 || !form.jobNumber || !form.date) ? 0.5 : 1 }}
                        disabled={sel.length === 0 || !form.jobNumber || !form.date}
                        onClick={() => submitCheckIn(inc)}
                      >
                        {t.incomingCheckInBtn(sel.length)}
                      </button>
                    </>
                  ) : (
                    <div className="text-sm" style={{ color: colors.green }}>{t.incomingFullyCheckedIn}</div>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// Reads a scanned Farspeed job sheet. The Excel originals are parsed properly by
// guessFieldsFromWorkbook; a PDF has no cells to read, so the page is scanned and the same
// structures are rebuilt here from what comes back. Everything downstream - the referral
// blocks, the case markings, the declared figures - is then identical to an Excel upload.
// One way in for both PDF scans. Two things were wrong before: the legacy job-sheet scan
// posted straight to api.anthropic.com, which has no key attached outside the preview, and
// both paths called response.json() on whatever came back. A proxy that times out returns
// a plain-text page, so a failed scan surfaced as "Unexpected token 'A'" rather than as
// what actually happened - and a 19-page packing list is exactly the kind of document that
// takes long enough to time one out.
// Reassembles a streamed reply. The stream is server-sent events: one "data:" line per
// chunk, each a small JSON object. Only the text deltas matter here; the stop reason is kept
// so that a reply which ran out of room can be reported as that rather than as bad JSON.
async function readPdfScanStream(response) {
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let text = "";
  let stopReason = "";
  let streamError = "";
  const handle = (payload) => {
    if (payload === "[DONE]") return;
    let evt;
    try { evt = JSON.parse(payload); } catch (err) { return; }
    if (evt.type === "content_block_delta" && evt.delta && typeof evt.delta.text === "string") {
      text += evt.delta.text;
    } else if (evt.type === "message_delta" && evt.delta && evt.delta.stop_reason) {
      stopReason = evt.delta.stop_reason;
    } else if (evt.type === "error") {
      streamError = (evt.error && (evt.error.message || evt.error)) || "stream error";
    }
  };
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    // Events are separated by a blank line; anything after the last one is a partial event
    // and stays in the buffer until the rest of it arrives.
    let sep;
    while ((sep = buffer.indexOf("\n\n")) !== -1) {
      const chunk = buffer.slice(0, sep);
      buffer = buffer.slice(sep + 2);
      for (const line of chunk.split("\n")) {
        if (line.startsWith("data:")) handle(line.slice(5).trim());
      }
    }
  }
  for (const line of buffer.split("\n")) {
    if (line.startsWith("data:")) handle(line.slice(5).trim());
  }
  if (streamError) throw new Error(streamError);
  const clean = text.replace(/```json|```/g, "").trim();
  if (!clean) throw new Error("empty reply from the scanner");
  try {
    return JSON.parse(clean);
  } catch (parseErr) {
    throw new Error(stopReason === "max_tokens" ? "reply-too-long" : "truncated-or-invalid-json");
  }
}
// The words the scanner is given. Lifted out of the one screen that used to own it so
// the packing list reader asks for exactly the same reading.
const PDF_SCAN_PROMPT = `This is a packing list, delivery memo, shipping list, or similar logistics document for elevator/escalator materials, possibly in English, Traditional or Simplified Chinese, or mixed.

Follow these extraction rules exactly - they keep the output compact even for long, dense documents:

1. Each row/block of the table is ONE case/package. For each case extract exactly these 5 things:
   - "code": the case/box number exactly as printed, letters and dashes included - "01A2101", "02C3102-3-1", "16D5416C", "C01". Never renumber or simplify them.
   - "description" on the group: a short generic name for what the order holds, e.g. "ELEVATOR MATERIALS" or the commodity named on an accompanying delivery order ("Guide Rail"). Used for any case whose own line gives no description.
   - "lot": for a Mitsubishi document, the Delivery Memo number - "13-DM-26-0500", printed as "DM No." - which is what these lots are filed under; fall back to the shipping mark (S/M) if the DM number is blank. Otherwise the lift/unit number. This is very often printed right next to the case number in parentheses, like "(#.01)" or "(#.23)" - extract just the number (e.g. "01", "23"). Different cases sharing the same case-number prefix (e.g. multiple "C21" cases) but different lift numbers are DIFFERENT packages in DIFFERENT lots. If there's no such lift marker anywhere, use the shop order number or another batch identifier as the lot instead.
   - "description": ONLY the general category/heading text for that case (e.g. "Guide Rail", "Rail Bracket", "Traction Machine", "Installation Material"). Do NOT list the individual part numbers or sub-components underneath it even if the document itemizes many - this is the single most important rule for keeping output size manageable on long documents.
   - "weightKg": use the GROSS weight (毛重 / GROSS column), not net weight - gross is what matters here.
   - "cbm": ONLY if a CBM/volume/m3 figure is printed directly for that case. If the document does not give one, use '' - do NOT try to work it out yourself.
   - "length", "width", "height": that case's dimensions exactly as printed, as plain numbers. These are very often three separate columns headed "Length cm", "Width cm", "Height cm" (or L/W/H, 長/闊/高), but may instead be one combined cell like "500*20*20" or "500x20x20" - read either form. Use '' for any you cannot find. Do not convert or round them.
   - "dimUnit": the unit those dimensions are printed in - "cm" or "mm". Take it from the column heading or a nearby note. If nothing says, use "" rather than guessing.
2. Group all cases by their lot/lift number into the "groups" array - one entry per distinct lot.
2b. Also look for shipping/bill-of-lading details anywhere on the document: vessel/ship name (often after "ex ss." or 船名), voyage number (航次), container numbers (貨櫃號), and bill-of-lading number (提單編號 / BL NO.). Combine them into one line for "ssDoNo" in roughly this style: ex ss."SHIP NAME" V.VOYAGE; CONTAINERS NO. XXXX/40GP. If none present, use ''. Also return them separately in "shipping".
2b-ii. IMPORTANT - a factory packing list often has a "package"/"Packages"/件數 column that gives the NUMBER of packages on that line, not a case number. One line reading "29 | APK00171P001 | T89/B" is twenty-nine packages, not one. Never count table rows as packages. For a document like that, return each line of the order in "lines" and leave "packages" empty:
   - "packages": how many packages that line covers, as a number. A dash or blank means zero - loose hardware travelling inside another line's cases.
   - "description": the short material description on that line.
   - "netWeightKg" and "grossWeightKg": both if both columns exist, else whichever is given. Do NOT pick between them; the app takes the heavier.
   - "cbm": only if that line states one.
   Group the lines by the order the document groups them under - "Order No.CED-1831" and the reference beside it - and put that whole heading in "lot", e.g. "CED-1831 (EL-1926)".
2b-iii. IMPORTANT - a single order is often subdivided further by a "Group" column carrying A, B, C, D against runs of lines, with a handwritten or printed annotation in the left margin beside each group naming what it is for: an INS reference like "0/24/576", a project number like "EL-1924" or "EL-1876", and the lifts like "#L2-L4", "#L-C01, L-C02, #L-C03, L-C04, #L-C05" or "#L-C06". Those groups go to different lifts and must NOT be merged.
   Return ONE entry in "groups" per group letter, each holding only that group's lines, and build its "lot" from the order, the group letter and the annotation, e.g. "CED-1833/B (EL-1876 #L-C01 to L-C05)". Put the group letter in "group" and the INS reference in "insRef". Read the margin annotations even when handwritten; where a group has none, use just the order and letter.
2b-iv. IMPORTANT - a Fujitec packing list has two levels of row and only the outer one is a package. An outer row carries a C/NO. (the case number), a PK NO. like "ZDZ1703+K-05A", N.W(KGS), G.W(KGS), Volume(M3) and a TYPE such as WOODEN CASE. Underneath it sit indented ITEM NO. rows listing that case's contents - PART NAME, PART NO., Job No., QTY PCS. Those inner rows are contents, NOT packages: never count them, and never take their weights.
   The lot is the job number at the FRONT of the PK NO., before the "+": "ZDZ1703+K-05A" and "ZDZ1703+Z-12HA" are both lot ZDZ1703. Use exactly that, e.g. "ZDZ1703".
   CRITICAL - do NOT use the inner rows' "Job No." column to decide which lot a case belongs to. One case can hold parts for several jobs: a case whose PK NO. begins "ZDZ1703" may list contents against both ZDZ1703 and ZDZ1708, but the case is ZDZ1703's. The PK NO. decides, always.
   Set "code" to the C/NO. exactly as printed ("01", "06", "32"), the weight from G.W(KGS), and the volume from Volume(M3).
   CRITICAL - each job is numbered separately and both usually start at 02, so the two jobs share most of their case numbers and differ only in their last few: one job may end 34-35 while the other ends 32-33. Read every C/NO. from the left-hand column of that case's own row and never carry a number over from the other job's section, never continue a job's numbering past what is printed, and never assume the two jobs end on the same numbers.
2b-v. A Shipping Marks block at the end of a packing list states, per job, which case numbers belong to it - "FUJITEC / ZDZ1703 / PO NO.HE-6717 / C/NO. 02-05, 09-10, 13-30, 34-35". Return every such block in "shippingMarks", copying the C/NO. line verbatim, ranges and all, with the job/order number from the mark as its "lot". Use it to check your grouping; where it disagrees with what you read off the rows, follow the Shipping Marks.
2c. IMPORTANT - many documents are NOT per-case tables at all. A Delivery Memo (DM), an arrival/release notice (到貨通知提貨單), or a shipping order states only the OVERALL totals - "29 Package(s)", "14.088 CBM", "12,909 Kgs", "29 件" - and then lists the case markings separately under a heading like "C/S NO." or "SHIPPING MARK", one marking per line, sometimes several comma-separated per line (e.g. "01C01,01C02,01C03"). For a document like that:
   - put the stated totals in "statedPackages", "statedWeightKg" and "statedCbm" on the group;
   - put every case marking, expanded from any comma-separated lines into individual entries, into "caseNumbers" on the group, exactly as printed;
   - leave "packages" as an empty array. Do NOT invent per-case weights or volumes for these - the totals are all the document states.
2c-ii. The list of markings is often on a separate attached page laid out in several columns across the page, each row holding the marking split into parts - a lift number, a component code and a case suffix, e.g. "09  B11  09" or "19  D41  21-2-2". Read each column from top to bottom, then move to the next column to its right. Every row is ONE package. Join that row's parts, in the order printed and with no spaces, to form the marking - "09B1109", "19D4121-2-2" - following the style the memo's own face uses where it prints a few of them.
2c-ii-b. On a packing list the case number and its lift marker sit on the SAME row, with the marker in parentheses: "B11 01  (#.01)", "E21 23  (#.23)", "B31 02 (#.02)". The lift number goes at the FRONT of the marking, not the back, and the row's own parts follow it in the order printed, with no spaces anywhere:
    "B11 01   (#.01)"  ->  "01B1101"
    "B11 02   (#.02)"  ->  "02B1102"
    "E21 23   (#.23)"  ->  "23E2123"
    "D11 01-3-1 (#.01)" ->  "01D1101-3-1"
  Never leave the space in "B11 01", never return the marking without its leading lift number, and never put the lift number on the end. This is the form every delivery memo, job sheet and depot record uses, and a marking built any other way matches nothing.
  Moving the lift number to the front of a marking does NOT make it the lot. The lot stays whatever rule 2c gives - the DM number for a Mitsubishi document, or the shipping mark or shop order where there is no DM - and the lift number belongs in front of each case marking and nowhere else. Returning "23" as the lot because the cases are lift 23 loses the reference every other document keys on.
  The lift number is READ from the document, never assumed: a job covering lifts 01 and 02 has only 01 and 02 in front, and one covering 23 and 24 has only those. Take it from that row's own "(#.nn)" marker, and where a row has none, from the lift the section it sits under is announced with. Do not carry a number over from an example or from another document.
2c-ii-c. Where a list of markings is separated by "&" before the last one - "01B81, 01D41, 01Z11 & 92B11" - the "&" is simply the final comma. Return that last marking like any other.
2c-iii. CRITICAL - return exactly the markings that are printed, and no others. If they come to fewer than the stated package count, return the ones you can see: do NOT invent, pad, repeat or renumber markings to reach the stated total, and do NOT drop any to match a smaller one. The document disagreeing with itself is something the reader needs told, and it is reported from the two figures.
2c-iv. Where a totals-only document's markings fall into more than one group, its face totals cover the whole document, so put them in "documentTotals" and NOT in "statedPackages"/"statedWeightKg"/"statedCbm" on any single group. Where everything is one group, put them on that group. Either way, state them once and never divide them yourself.
2c-v. A packing list is often sent with a delivery order or arrival notice covering the same shipment. Where the total CBM or weight appears only on that companion page and not per order, put it in "documentTotals" - do NOT divide it between the orders yourself, and do NOT copy it onto one of them.
2d. Look for terminal/storage dates: the arrival/ETA date (到港日期 / ETA) as "terminalArrivalDate" and the last free storage day (免費倉期 ... 至) as "lastFreeDay", both as YYYY-MM-DD. Use '' if absent.
3. Keep everything as compact as possible: short descriptions, no commentary, no repeated sub-item lists.

Respond with ONLY a raw JSON object in EXACTLY this shape and nothing else (no markdown fences, no commentary, no explanation before or after):
{"client": "best-guess client name or ''", "project": "site/building/project name found in the document, or ''", "ssDoNo": "vessel + voyage + container line or ''", "shipping": {"vessel": "", "voyage": "", "blNo": "", "containerNo": ""}, "terminalArrivalDate": "YYYY-MM-DD or ''", "lastFreeDay": "YYYY-MM-DD or ''", "documentTotals": {"packages": number_or_empty_string, "weightKg": number_or_empty_string, "cbm": number_or_empty_string}, "shippingMarks": [{"lot": "job/order number in the mark", "cases": "that mark's C/NO. line exactly as printed, ranges and all"}], "groups": [{"lot": "lift/lot/shop-order number identifying this batch", "containers": ["container numbers if any, else empty array"], "statedPackages": number_or_empty_string, "statedWeightKg": number_or_empty_string, "statedCbm": number_or_empty_string, "caseNumbers": ["case markings, one per entry, only for totals-only documents"], "group": "group letter if the document has one, else ''", "insRef": "INS/works reference beside the group, e.g. 0/24/576, else ''", "lines": [{"packages": number, "description": "", "netWeightKg": number_or_empty_string, "grossWeightKg": number_or_empty_string, "cbm": number_or_empty_string}], "packages": [{"code": "case/package number", "description": "short category name, a few words only", "weightKg": number_or_empty_string, "cbm": number_or_empty_string, "length": number_or_empty_string, "width": number_or_empty_string, "height": number_or_empty_string, "dimUnit": "cm_or_mm_or_empty"}]}]}
If the document only has one overall lot/shipment with no explicit lift/case breakdown, put everything under a single group with a sensible lot name.

SIZE MATTERS - a long document has to be answered before the request times out, so keep the reply as short as it can be while staying complete. OMIT any key you would set to '' or to an empty array: write {"code":"01","weightKg":17,"cbm":0.08} rather than repeating every field with empty values. Never abbreviate or omit a case itself - every package must appear - but say nothing about its contents, and use no whitespace beyond what JSON requires.`;
async function postPdfScan(body, attempt = 0) {
  let response;
  try {
    response = await fetch("/api/scan-pdf", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  } catch (netErr) {
    // A dropped connection on a long document is often transient; one retry is worth more
    // than making someone re-pick the file.
    if (attempt === 0) return postPdfScan(body, 1);
    throw new Error(`network: ${netErr && netErr.message ? netErr.message : netErr}`);
  }
  if ((response.status >= 500 || response.status === 408) && attempt === 0) {
    return postPdfScan(body, 1);
  }
  // The proxy streams the reply so that it starts arriving within the edge function's
  // time-to-first-byte limit. Plain JSON is still accepted, both for errors (which are never
  // streamed) and for a proxy that has not been updated yet.
  const contentType = (response.headers && response.headers.get("content-type")) || "";
  if (contentType.includes("text/event-stream") && response.body) {
    return readPdfScanStream(response);
  }
  const raw = await response.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch (parseErr) {
    // Not JSON at all: report the status and the first of the body, which is usually the
    // one line that says what went wrong.
    const snippet = raw.replace(/\s+/g, " ").trim().slice(0, 120);
    throw new Error(`server returned ${response.status}: ${snippet || "empty response"}`);
  }
  if (!response.ok || data.error) {
    throw new Error((data.error && (data.error.message || data.error)) || `server returned ${response.status}`);
  }
  const text = (data.content || []).map((b) => b.text || "").join("");
  const clean = text.replace(/```json|```/g, "").trim();
  try {
    return JSON.parse(clean);
  } catch (parseErr) {
    // A reply that ran out of room mid-object is the usual cause here.
    throw new Error(data.stop_reason === "max_tokens" ? "reply-too-long" : "truncated-or-invalid-json");
  }
}
async function scanLegacyJobSheetPdf(file) {
  const base64 = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(",")[1]);
    r.onerror = () => rej(new Error("read failed"));
    r.readAsDataURL(file);
  });
  const prompt = `This is a scanned Farspeed Contractors job sheet (工單 / JOB SHEET) for a Hong Kong lift depot. Read it and return JSON only - no prose, no markdown fences.

The header carries: ACCOUNT (客戶) the client; TO (送) the site, often an English line and a Chinese line; JOB NO. (快達單號); DATE (日期); ORDERED BY (落單人); JOB REF. (地盤代號); P.O. NO.; SS/D.O. NO. (提單資料).

The body is one or more referral blocks. Each begins "Ref Job no. NNNNNNN on DD/MM/YYYY" (sometimes "Refer to job no."), and contains a "C/S No." line stating that block's PKGS, KGS and CBM, followed by that block's case numbers. Case numbers are printed exactly as they are used - "01A11", "02D5102A", "01C3101-4-1", or plain numbers like "1-3/3" - listed comma-separated and wrapped over as many lines as needed. A "LIFT NO." line above the blocks gives the lift(s). A handwritten or printed "Refer. DM no. 13-DM-26-NNNN" may appear; return it as dmNo. A final "共:" line gives the sheet total.

Return exactly this shape:
{"docType":"Delivery or CFS or Devan or Shifting or Hoisting","client":"","projectEn":"","projectZh":"","jobNumber":"","date":"YYYY-MM-DD","jobRef":"","orderedBy":"","poNo":"","ssDoNo":"","shkNumber":"","dmNo":"","liftNo":"","blocks":[{"refJobNumber":"","refDate":"YYYY-MM-DD","liftNo":"","dmNo":"","shippingMark":"","pkgs":"","kg":"","cbm":"","caseNumbers":["every case number of this block, one per entry, exactly as printed"]}],"total":{"pkgs":"","kg":"","cbm":""}}

Rules: numbers as plain digits with no thousands separators. Dates as YYYY-MM-DD; the sheet writes them day-first. Use '' for anything not present. Never invent a case number and never renumber one - if a block lists no cases, return an empty array.`;
  return postPdfScan({
    model: "claude-sonnet-4-6",
    max_tokens: 8000,
    messages: [{
      role: "user",
      content: [
        { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
        { type: "text", text: prompt },
      ],
    }],
  });
}
// Rebuilds the row fields a workbook upload would have produced.
function legacyFieldsFromScan(parsed) {
  const out = {
    client: parsed.client || "", projectEn: parsed.projectEn || "", projectZh: parsed.projectZh || "",
    jobNumber: parsed.jobNumber || "", date: parsed.date || "", jobRef: parsed.jobRef || "",
    ssDoNo: parsed.ssDoNo || "", shkNumber: parsed.shkNumber || "",
    packageCount: (parsed.total && parsed.total.pkgs) || "",
    weightKg: (parsed.total && parsed.total.kg) || "",
    volumeCbm: (parsed.total && parsed.total.cbm) || "",
    unitCode: parsed.liftNo || "",
    refBlocks: [], declaredTotalsList: [], caseMarksByLot: {}, caseMarksByRef: {},
    caseCodesByLot: {}, caseCountMismatches: [],
  };
  const blocks = parsed.blocks || [];
  out.referJobNumber = [...new Set(blocks.map((b) => String(b.refJobNumber || "").trim()).filter(Boolean))].join(", ");
  out.referDate = (blocks.find((b) => b.refDate) || {}).refDate || "";
  for (const b of blocks) {
    const codes = (b.caseNumbers || []).map((c) => String(c || "").trim()).filter(Boolean);
    // A block whose cases are plain numbers keeps the numeric reading, so "1-3/3" still
    // behaves like a marking off an Excel sheet; anything with letters is a code.
    const numeric = codes.length > 0 && codes.every((c) => /^\d+(\/\d+)?$/.test(c));
    const lotRef = String(b.dmNo || parsed.dmNo || b.shippingMark || b.refJobNumber || "").trim();
    const unitCode = String(b.liftNo || parsed.liftNo || "").trim();
    const lot = {
      lotRef, altRef: String(b.shippingMark || "").trim(), unitCode,
      caseNumbers: numeric ? codes.map((c) => Number(String(c).split("/")[0])) : [],
      caseCodes: numeric ? [] : codes,
      caseText: codes.join(", "), lotCases: null,
      pkgs: String(b.pkgs || ""), kg: String(b.kg || ""), cbm: String(b.cbm || ""),
      shkNumber: parsed.shkNumber || "",
    };
    out.refBlocks.push({
      refJobNumber: String(b.refJobNumber || "").trim(), refDate: b.refDate || "",
      shkNumber: parsed.shkNumber || "", liftNo: unitCode, lots: [lot],
      pkgs: lot.pkgs, kg: lot.kg, cbm: lot.cbm, caseNumbers: lot.caseNumbers,
    });
    const stated = Number(lot.pkgs);
    if (stated > 0 && codes.length > 0 && stated !== codes.length) {
      out.caseCountMismatches.push({ lot: lotRef || unitCode || "", stated, listed: codes.length });
    }
    if (lot.kg || lot.cbm) {
      out.declaredTotalsList.push({
        pkgs: lot.pkgs, kg: lot.kg, cbm: lot.cbm,
        refJobNumber: lot.lotRef ? String(b.refJobNumber || "") : "", shkNumber: parsed.shkNumber || "",
        context: [b.refJobNumber, parsed.shkNumber, unitCode, lotRef, lot.altRef].filter(Boolean).join(" "),
      });
    }
    for (const key of [lotRef, lot.altRef, unitCode, String(b.refJobNumber || "").trim()]) {
      if (!key) continue;
      const k = key.toUpperCase();
      if (lot.caseCodes.length) {
        out.caseCodesByLot[k] = {
          codes: [...new Set([...((out.caseCodesByLot[k] || {}).codes || []), ...lot.caseCodes])],
          text: lot.caseText,
        };
      } else if (lot.caseNumbers.length) {
        out.caseMarksByLot[k] = {
          numbers: [...new Set([...((out.caseMarksByLot[k] || {}).numbers || []), ...lot.caseNumbers])].sort((x, y) => x - y),
          text: lot.caseText, lotCases: null,
        };
      }
    }
  }
  return out;
}
function LegacyUploadsPanel({ onReplaceIncomingCases, employees, setDirectory, legacyArchive, setLegacyArchive, items, incoming, onLegacyCheckIn, onLegacyCheckInBatch, directory, onLegacyImport, onLegacyDeliver, onLegacyEnrich, onLegacyReverse, onAddIncoming, colors, t, lang }) {
  const [rows, setRows] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState(null);
  const [backlogSearch, setBacklogSearch] = useState("");
  const [backlogJobFilter, setBacklogJobFilter] = useState("All");
  const [backlogSiteFilter, setBacklogSiteFilter] = useState("All");
  const [backlogJobRefFilter, setBacklogJobRefFilter] = useState("All");
  const [backlogClientFilter, setBacklogClientFilter] = useState("All");
  const [backlogFileFilter, setBacklogFileFilter] = useState("All");
  const [backlogSort, setBacklogSort] = useState("recent");
  const [backlogTypeFilter, setBacklogTypeFilter] = useState("All");
  const [picked, setPicked] = useState({});
  // Deleting a site's backlog one row at a time is a hundred clicks, and a rebuild starts by
  // clearing it. Selection follows the filters, so "select all" means all the rows on screen
  // and never the ones a filter is hiding - which is the only way this is safe to offer.
  const pickedIds = Object.keys(picked).filter((id) => picked[id]);
  function deletePicked(alsoReverse) {
    const rows = (legacyArchive || []).filter((r) => picked[r.id]);
    if (!rows.length) return;
    const plans = alsoReverse ? rows.map((r) => reversalPlanFor(r)) : [];
    const entries = plans.reduce((n, p) => n + p.items.length, 0);
    const removed = plans.reduce((n, p) => n + p.items.filter((x) => x.remove).length, 0);
    if (!window.confirm(alsoReverse
      ? t.legacyBulkReverseConfirm(rows.length, entries, removed)
      : t.legacyBulkDeleteConfirm(rows.length))) return;
    if (alsoReverse) plans.forEach((p) => onLegacyReverse && onLegacyReverse(p));
    const ids = new Set(rows.map((r) => r.id));
    ids.forEach((id) => storageSet(`legacyDoc:${id}`, "").catch(() => {}));
    setLegacyArchive((prev) => prev.filter((r) => !ids.has(r.id)));
    setPicked({});
  }
  const [editingBacklogId, setEditingBacklogId] = useState(null);
  const [deletingBacklogId, setDeletingBacklogId] = useState(null);
  const [deletePlan, setDeletePlan] = useState(null);
  // Every record a given archived file created, so it can be corrected from the archive
  // listing rather than hunted down entry by entry in Inventory. A legacy check-in stamps
  // its arrival batch with the file it came from, and a legacy delivery puts the file name
  // in its notes, so each file can find its own records back. Date is the fallback for
  // anything recorded before that stamping existed.
  function linkedRecordsFor(r) {
    const ids = (r.linkedItemIds && r.linkedItemIds.length)
      ? r.linkedItemIds
      : String(r.linkedItemId || "").split(",").map((s) => s.trim()).filter(Boolean);
    const out = [];
    for (const id of ids) {
      const it = (items || []).find((i) => i.id === id);
      if (!it) continue;
      if (r.docType === "Delivery") {
        const live = (it.deliveries || []).filter((d) => !d.cancelled);
        const rec = live.find((d) => r.fileName && String(d.notes || "").includes(r.fileName))
          || live.find((d) => r.jobNumber && d.jobNumber === r.jobNumber);
        if (rec) out.push({ item: it, kind: "delivery", rec });
      } else {
        const rec = (it.arrivals || []).find((a) => r.fileName && a.declaredSource === r.fileName)
          || (it.arrivals || []).find((a) => r.date && a.date === r.date);
        if (rec) out.push({ item: it, kind: "arrival", rec });
      }
    }
    return out;
  }
  function draftRecordsFor(r) {
    return linkedRecordsFor(r).map(({ item, kind, rec }) => ({
      itemId: item.id,
      // Where the record sits today, kept apart from itemId so the picker can point it at a
      // different entry and the save still knows which one to lift it off.
      originalItemId: item.id,
      kind, recId: rec.id,
      codes: rec.codes || [],
      date: rec.date || "",
      type: rec.type || ARRIVING_TYPES[0],
      jobNumber: rec.jobNumber || "",
      kg: (rec.declared && rec.declared.kg) || "",
      cbm: (rec.declared && rec.declared.cbm) || "",
    }));
  }
  // Entries a record can be moved to. Same client as the archive listing, since a delivery
  // never crosses between clients, plus whichever entry the record is on now so the picker
  // always has its own value to show.
  function moveTargetsFor(client, currentIds) {
    const keep = new Set(currentIds || []);
    return (items || [])
      .filter((i) => keep.has(i.id) || (!i.cancelled && (!client || i.client === client)))
      .sort((a, b) => String(a.id).localeCompare(String(b.id)));
  }
  function itemLabel(item) {
    if (!item) return "";
    return `${item.id}${item.unitCode ? ` \u00b7 ${item.unitCode}` : ""}`;
  }
  // How many of a delivery's case numbers exist on the entry it is being moved to. A lift
  // and its twin are cased up the same way, so most codes carry across, but any that do not
  // would point at cases the new entry does not hold and are dropped on the move.
  function codesCarriedOver(codes, targetId) {
    const target = (items || []).find((i) => i.id === targetId);
    const have = new Set(((target && target.packages) || []).map((p) => String(p.code || "").trim()));
    return (codes || []).filter((c) => have.has(String(c || "").trim()));
  }
  // Writes the corrected records back onto their entries. A figure typed here is stated
  // for this lot, so it drops the "share of a bigger total" flag that would otherwise mark
  // it an estimate and see it overruled by per-case packing-list weights.
  //
  // A record can also be moved to a different entry, which is how a delivery filed against
  // the wrong lift gets put right: it is lifted off the entry it is on and appended to the
  // chosen one. Every entry the save touches is worked on as a single copy and patched once
  // at the end - patching per record would let two records landing on the same entry
  // overwrite each other, since the enrich handler keys its patches by entry.
  function saveLinkedRecords(records) {
    if (!onLegacyEnrich || !records || !records.length) return [];
    const work = new Map();
    const load = (id) => {
      if (!id) return null;
      if (!work.has(id)) {
        const it = (items || []).find((i) => i.id === id);
        if (!it) return null;
        work.set(id, {
          base: it,
          arrivals: [...(it.arrivals || [])],
          deliveries: [...(it.deliveries || [])],
          arrivalsDirty: false, deliveriesDirty: false,
        });
      }
      return work.get(id);
    };
    for (const rec of records) {
      const from = load(rec.originalItemId || rec.itemId);
      if (!from) continue;
      const isArrival = rec.kind === "arrival";
      const list = isArrival ? from.arrivals : from.deliveries;
      const idx = list.findIndex((x) => x.id === rec.recId);
      if (idx === -1) continue;
      const declared = (String(rec.kg).trim() || String(rec.cbm).trim())
        ? { kg: String(rec.kg).trim(), cbm: String(rec.cbm).trim(), split: false } : null;
      let next = isArrival
        ? { ...list[idx], date: rec.date, type: rec.type, declared: declared && { ...(list[idx].declared || {}), ...declared }, declaredEdited: true }
        : { ...list[idx], date: rec.date, jobNumber: rec.jobNumber, declared };
      const target = rec.itemId && rec.itemId !== (rec.originalItemId || rec.itemId) ? load(rec.itemId) : null;
      if (!target) {
        list[idx] = next;
        if (isArrival) from.arrivalsDirty = true; else from.deliveriesDirty = true;
        continue;
      }
      list.splice(idx, 1);
      if (isArrival) from.arrivalsDirty = true; else from.deliveriesDirty = true;
      if (!isArrival && (next.codes || []).length) {
        next = { ...next, codes: codesCarriedOver(next.codes, rec.itemId) };
      }
      if (isArrival) { target.arrivals.push(next); target.arrivalsDirty = true; }
      else { target.deliveries.push(next); target.deliveriesDirty = true; }
    }
    const entries = [];
    for (const [id, w] of work) {
      const patch = {};
      // Only what actually changed goes into the patch: an entry carrying arrivals is
      // recomputed from them, and recomputing one whose weight was corrected by hand would
      // throw that correction away.
      if (w.arrivalsDirty) {
        patch.arrivals = w.arrivals;
        const dates = w.arrivals.map((a) => a.date).filter(Boolean).sort();
        patch.depotArrivalDate = dates[0] || w.base.depotArrivalDate;
      }
      if (w.deliveriesDirty) patch.deliveries = w.deliveries;
      if (Object.keys(patch).length) entries.push({ itemId: id, patch });
    }
    if (entries.length) onLegacyEnrich(entries);
    return entries.map((e) => e.itemId);
  }
  // Everything a file put into the depot, and what undoing it would mean, worked out before
  // anything is touched so it can be shown and agreed to first.
  //
  // A Devan or CFS arrival is lifted off its entry and its cases go back to the Incoming
  // shipment they were checked in from, leaving that shipment as a packing list waiting to
  // be checked in - which is where it stood before the file was processed. A delivery is
  // lifted off and its cases return to store.
  //
  // An entry is only removed when this file is what created it and nothing else is left on
  // it. Being empty is not enough on its own: a legacy-imported entry carries no arrival
  // records at all, so a delivery file that happened to be the only thing on one would look
  // just as empty while being nothing of the sort.
  function reversalPlanFor(r) {
    const ids = (r.linkedItemIds && r.linkedItemIds.length)
      ? r.linkedItemIds
      : String(r.linkedItemId || "").split(",").map((s) => s.trim()).filter(Boolean);
    const byItem = new Map();
    const seed = (item) => {
      if (!byItem.has(item.id)) byItem.set(item.id, { itemId: item.id, item, arrivalIds: [], deliveryIds: [], codes: [] });
      return byItem.get(item.id);
    };
    for (const id of ids) {
      const it = (items || []).find((i) => i.id === id);
      if (it) seed(it);
    }
    for (const { item, kind, rec } of linkedRecordsFor(r)) {
      const p = seed(item);
      if (kind === "arrival") p.arrivalIds.push(rec.id); else p.deliveryIds.push(rec.id);
      p.codes = [...new Set([...p.codes, ...((rec.codes) || [])])];
    }
    const plan = { fileName: r.fileName, docType: r.docType, items: [], incoming: [], stranded: [], blocked: [] };
    for (const p of byItem.values()) {
      const arrivalsLeft = (p.item.arrivals || []).filter((a) => !p.arrivalIds.includes(a.id));
      const deliveriesLeft = (p.item.deliveries || []).filter((d) => !p.deliveryIds.includes(d.id) && !d.cancelled);
      const createdHere = String(p.item.notes || "").includes(r.fileName)
        || (p.arrivalIds.length > 0 && arrivalsLeft.length === 0);
      p.remove = r.docType !== "Delivery" && createdHere && arrivalsLeft.length === 0 && deliveriesLeft.length === 0;
      p.label = itemLabel(p.item);
      p.arrivalsLeft = arrivalsLeft.length;
      p.deliveriesLeft = deliveriesLeft.length;
      p.caseCount = p.codes.length;
      if (!p.arrivalIds.length && !p.deliveryIds.length && !p.remove) {
        // Linked, but nothing of this file's making is still on it - already reversed by
        // hand, or moved elsewhere. Said out loud rather than silently skipped. An entry
        // this file did create but which has since been delivered off is a different case
        // and gets its own line, since deleting it would take that delivery with it.
        if (createdHere && deliveriesLeft.length > 0) plan.blocked.push({ label: p.label, deliveries: deliveriesLeft.length });
        else plan.stranded.push(p.label);
        continue;
      }
      plan.items.push(p);
      if (p.arrivalIds.length) {
        for (const inc of (incoming || [])) {
          if (inc.linkedItemId !== p.item.id) continue;
          const codes = (p.item.arrivals || [])
            .filter((a) => p.arrivalIds.includes(a.id))
            .flatMap((a) => a.codes || []);
          if (codes.length) plan.incoming.push({ incomingId: inc.id, codes, unlink: p.remove });
        }
      }
    }
    return plan;
  }
  const [backlogEditDraft, setBacklogEditDraft] = useState(null);
  const fileInputRef = React.useRef(null);

  const [scanning, setScanning] = useState(false);

  async function handleFilesSelected(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    setScanning(true);
    const newRows = [];
    for (const file of files) {
      const base = {
        file,
        docType: guessDocTypeFromName(file.name),
        client: "",
        projectEn: "",
        projectZh: "",
        jobNumber: guessJobNumberFromName(file.name),
        date: "",
        unitCode: "",
        packageCount: "",
        weightKg: "",
        volumeCbm: "",
        ssDoNo: "",
        shkNumber: "",
        jobRef: "",
        oversizeByLot: {},
        referJobNumber: "",
        referDate: "",
        declaredTotalsList: [],
        refBlocks: [],
        caseMarksByLot: {},
        caseCodesByLot: {},
        caseCountMismatches: [],
        caseMarksByRef: {},
        caseAutoApplied: {},
        autoDetected: false,
      };
      const isExcel = /\.(xlsx|xls|csv)$/i.test(file.name);
      if (isExcel) {
        try {
          const buf = await file.arrayBuffer();
          // bookFiles keeps the workbook's raw parts, which is the only way to see that a
          // sheet's case list is a pasted picture rather than empty cells.
          const wb = XLSX.read(buf, { type: "array", cellDates: true, bookFiles: true });
          // A spreadsheet from the job sheet importer carries many job sheets at once, so
          // it is unpacked back into one staged row per sheet before anything else is
          // tried. From there it is treated exactly like a stack of uploaded job sheets:
          // matched to what the depot holds, cases pre-selected, and anything missing or
          // in the wrong place shown for correcting by hand before it is processed.
          const grid = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: false });
          const fromSpreadsheet = rowsFromJobSheetSpreadsheet(grid, file, resolveClientGuess);
          if (fromSpreadsheet && fromSpreadsheet.length) {
            newRows.push(...fromSpreadsheet);
            continue;
          }
          const guessed = guessFieldsFromWorkbook(wb, { docType: base.docType });
          const clientMatch = resolveClientGuess(guessed.client);
          Object.assign(base, {
            client: clientMatch || base.client,
            projectEn: guessed.projectEn || base.projectEn,
            projectZh: guessed.projectZh || base.projectZh,
            jobNumber: guessed.jobNumber || base.jobNumber,
            date: guessed.date || base.date,
            unitCode: guessed.unitCode || base.unitCode,
            packageCount: guessed.packageCount || base.packageCount,
            weightKg: guessed.weightKg || base.weightKg,
            volumeCbm: guessed.volumeCbm || base.volumeCbm,
            ssDoNo: guessed.ssDoNo || base.ssDoNo,
            jobRef: guessed.jobRef || base.jobRef,
            oversizeByLot: guessed.oversizeByLot && Object.keys(guessed.oversizeByLot).length ? guessed.oversizeByLot : base.oversizeByLot,
            referJobNumber: guessed.referJobNumber || base.referJobNumber,
            referDate: guessed.referDate || base.referDate,
            shkNumber: guessed.shkNumber || base.shkNumber,
            declaredTotalsList: guessed.declaredTotalsList || base.declaredTotalsList,
            refBlocks: guessed.refBlocks || base.refBlocks,
            caseMarksByLot: guessed.caseMarksByLot || base.caseMarksByLot,
            caseCodesByLot: guessed.caseCodesByLot || base.caseCodesByLot,
            caseCountMismatches: guessed.caseCountMismatches || [],
            caseMarksByRef: guessed.caseMarksByRef || base.caseMarksByRef,
            autoDetected: true,
          });
          if (base.docType === "Delivery" && guessed.referJobNumber) base.jobNumber = base.jobNumber || guessed.referJobNumber;
        } catch (err) { /* fall back to filename-only guesses */ }
      } else if (/\.pdf$/i.test(file.name)) {
        try {
          const parsed = await scanLegacyJobSheetPdf(file);
          const guessed = legacyFieldsFromScan(parsed);
          const clientMatch = resolveClientGuess(guessed.client);
          Object.assign(base, {
            docType: parsed.docType || base.docType,
            client: clientMatch || base.client,
            projectEn: guessed.projectEn || base.projectEn,
            projectZh: guessed.projectZh || base.projectZh,
            jobNumber: guessed.jobNumber || base.jobNumber,
            date: guessed.date || base.date,
            unitCode: guessed.unitCode || base.unitCode,
            packageCount: guessed.packageCount || base.packageCount,
            weightKg: guessed.weightKg || base.weightKg,
            volumeCbm: guessed.volumeCbm || base.volumeCbm,
            ssDoNo: guessed.ssDoNo || base.ssDoNo,
            jobRef: guessed.jobRef || base.jobRef,
            shkNumber: guessed.shkNumber || base.shkNumber,
            referJobNumber: guessed.referJobNumber || base.referJobNumber,
            referDate: guessed.referDate || base.referDate,
            declaredTotalsList: guessed.declaredTotalsList,
            refBlocks: guessed.refBlocks,
            caseMarksByLot: guessed.caseMarksByLot,
            caseCodesByLot: guessed.caseCodesByLot,
            caseCountMismatches: guessed.caseCountMismatches,
            hasPastedContentImage: guessed.hasPastedContentImage,
            autoDetected: true,
            scannedFromPdf: true,
          });
          if (base.docType === "Delivery" && guessed.referJobNumber) base.jobNumber = base.jobNumber || guessed.referJobNumber;
        } catch (err) {
          base.scanError = String(err && err.message ? err.message : err);
        }
      }
      newRows.push(base);
    }
    setRows((prev) => [...prev, ...newRows]);
    setResults(null);
    setScanning(false);
  }

  function updateRow(idx, next) {
    setRows((prev) => prev.map((r, i) => (i === idx ? next : r)));
  }
  function removeRow(idx) {
    setRows((prev) => prev.filter((_, i) => i !== idx));
  }

  async function processAll() {
    setProcessing(true);
    const archiveEntries = [];
    const importRows = [];
    const jobNoToImportIdx = new Map(); // arrival job number -> index into importRows, for same-batch linking
    const pendingEnrichments = [];
    const checkInOps = []; // {op, archiveEntry} pairs, resolved in one batch call after the loop
    const fileUriById = {};

    // Pass 1: archive every file, and create the arrival (Devan/CFS = add to storage).
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const id = `LEG${Date.now()}${Math.floor(Math.random() * 1000)}-${i}`;
      let fileUri = null;
      try {
        fileUri = await compressFileToDataUri(row.file);
      } catch (e) { /* archive without file if compression fails */ }
      if (fileUri) {
        try { await storageSet(`legacyDoc:${id}`, JSON.stringify({ uri: fileUri, name: legacySourceName(row), at: todayStr() })); } catch (e) {}
      }
      fileUriById[i] = fileUri;

      const hasSite = !!(row.projectEn || row.projectZh);
      let linkedItemIndex = null;
      let enrichTargetId = null;
      const matchedIncomings = (row.docType === "Devan" || row.docType === "CFS") && row.client && hasSite
        ? (incoming || []).filter((inc) => {
            if (inc.client !== row.client) return false;
            const done = new Set(inc.checkedInCodes || []);
            if ((inc.packages || []).filter((p) => !done.has(p.code)).length === 0) return false;
            return sitesLooselyMatch(row.projectEn, row.projectZh, inc.project, inc.constructionSite);
          })
        : [];
      const selectedByIncoming = row.selectedByIncoming || {};
      const incomingDeclaredDist = distributeDeclaredAcrossLots(
        row.declaredTotalsList,
        matchedIncomings,
        (inc) => sumSelectedPackages(inc.packages, selectedByIncoming[inc.id] || []),
        row.declaredByGroup || {}
      );
      const hasAnyIncomingSelection = matchedIncomings.some((inc) => (selectedByIncoming[inc.id] || []).length > 0);
      const archiveEntry = {
        id, rowIndex: i, fileName: legacySourceName(row), docType: row.docType, client: row.client,
        project: [row.projectEn, row.projectZh].filter(Boolean).join(" / "),
        jobNumber: row.jobNumber, jobRef: row.jobRef || "", date: row.date, uploadedAt: todayStr(), hasFile: !!fileUri,
        linkedItemId: null, linkedItemIds: [], __linkedItemIndex: null,
      };
      archiveEntries.push(archiveEntry);

      if (row.docType !== "Delivery" && row.client && hasSite) {
        if (hasAnyIncomingSelection && onLegacyCheckInBatch) {
          // Matched to one or more real Incoming shipments with cases selected - queue a
          // check-in for each one (creates or extends its linked inventory item, same as
          // the Incoming tab), resolved together with every other row's check-ins after
          // this loop.
          for (const inc of matchedIncomings) {
            const codes = selectedByIncoming[inc.id] || [];
            if (codes.length === 0) continue;
            const detectedOs = (row.oversizeByLot || {})[inc.unitCode];
            const oversize = normaliseOversize((row.oversizeByIncoming || {})[inc.id]
              || (detectedOs && detectedOs.length ? { checked: true, cases: detectedOs } : null));
            const declaredEdited = (row.declaredByIncoming || {})[inc.id];
            const declaredParsed = incomingDeclaredDist[inc.id];
            // Same precedence the row displayed: a lot inside a shared total is driven by
            // that total, and travels with the split flag so the item knows its weight is
            // a share rather than something the sheet stated for this lot.
            const declared = (declaredParsed && declaredParsed.split)
              ? { pkgs: declaredParsed.pkgs, kg: declaredParsed.kg, cbm: declaredParsed.cbm, split: true }
              : (declaredEdited || (declaredParsed
                  ? { pkgs: declaredParsed.pkgs, kg: declaredParsed.kg, cbm: declaredParsed.cbm }
                  : null));
            checkInOps.push({
              op: {
                incomingId: inc.id,
                codes,
                declared,
                declaredSource: legacySourceName(row),
                type: row.docType,
                depot: row.depot || DEPOTS[0],
                jobNumber: row.jobNumber,
                date: row.date,
                ssDoNo: row.ssDoNo,
                shkNumber: row.shkNumber,
                jobRef: row.jobRef,
                isOversize: !!(oversize.checked && cleanOversizeCases(oversize.cases).length),
                oversizeCases: oversize.checked ? cleanOversizeCases(oversize.cases) : [],
                oversizeCbm: oversize.checked ? String(oversizeCbmTotal(oversize.cases) || "") : "",
              },
              archiveEntry,
            });
          }
        } else {
        const jobNoTrim = String(row.jobNumber || "").trim();
        const existingByJobNo = jobNoTrim ? items.find((it) => String(it.jobNumber || "").trim() === jobNoTrim) : null;
        if (existingByJobNo) {
          // A real item already exists for this job number - most likely created via the
          // proper Packing List import, which has real per-case CBM/KG/case numbers.
          // Enrich it with this file's metadata rather than creating a duplicate that
          // would only carry flat totals and shadow the good per-case data.
          const patch = {};
          if (!existingByJobNo.ssDoNo && row.ssDoNo) patch.ssDoNo = row.ssDoNo;
          if (!existingByJobNo.shkNumber && row.shkNumber) patch.shkNumber = row.shkNumber;
          if (!existingByJobNo.jobRef && row.jobRef) patch.jobRef = row.jobRef;
          if (!existingByJobNo.project && row.projectEn) patch.project = row.projectEn;
          if (!existingByJobNo.constructionSite && row.projectZh) patch.constructionSite = row.projectZh;
          if (!existingByJobNo.depotArrivalDate && row.date) patch.depotArrivalDate = row.date;
          patch.notes = [existingByJobNo.notes, t.legacyEnrichedNote(legacySourceName(row))].filter(Boolean).join(" \u00b7 ");
          enrichTargetId = existingByJobNo.id;
          pendingEnrichments.push({ itemId: existingByJobNo.id, patch });
        } else {
          const dirMatch = (directory || []).find((s) =>
            (row.projectEn && s.siteEn && s.siteEn.toLowerCase() === row.projectEn.toLowerCase()) ||
            (row.projectZh && s.siteZh && s.siteZh === row.projectZh)
          );
          const newItem = {
            ...emptyForm(),
            client: row.client,
            project: row.projectEn || (dirMatch ? dirMatch.siteEn : ""),
            constructionSite: row.projectZh || (dirMatch ? dirMatch.siteZh : ""),
            jobNumber: row.jobNumber,
            jobRef: row.jobRef || "",
            depotArrivalDate: row.date,
            unitCode: row.unitCode,
            packageCount: row.packageCount || "",
            weightKg: row.weightKg || "",
            volumeCbm: row.volumeCbm || "",
            arrivingType: row.docType === "CFS" ? "CFS" : "Devan",
            ssDoNo: row.ssDoNo || "",
            shkNumber: row.shkNumber || "",
            notes: t.legacyImportedNote(legacySourceName(row)),
            deliveries: [],
          };
          linkedItemIndex = importRows.length;
          if (row.jobNumber) jobNoToImportIdx.set(String(row.jobNumber).trim(), linkedItemIndex);
          importRows.push(newItem);
        }
        }
      }
      archiveEntry.linkedItemId = enrichTargetId;
      archiveEntry.__linkedItemIndex = linkedItemIndex;
    }

    if (checkInOps.length > 0 && onLegacyCheckInBatch) {
      const results = onLegacyCheckInBatch(checkInOps.map((c) => c.op));
      results.forEach((itemId, idx) => {
        if (!itemId) return;
        const entry = checkInOps[idx].archiveEntry;
        if (!entry.linkedItemIds.includes(itemId)) entry.linkedItemIds.push(itemId);
        entry.linkedItemId = entry.linkedItemIds.join(", ");
      });
    }

    // Pass 2: Delivery = subtract from storage. Rows matched to real inventory items (via
    // the case-selection UI) deliver those exact codes; rows with no match fall back to
    // the old referJobNumber + flat package count behaviour.
    const existingDeliveryEntries = [];
    const returnEntries = [];
    let sameBatchDeliveryCount = 0;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      // A Return runs the same matching and the same case selection as a Delivery, but the
      // cases are coming back rather than going out. Handled below by cancelling the
      // delivery records that took them out, which is what physically happened: the goods
      // returned, so they were never really delivered.
      if (row.docType !== "Delivery" && row.docType !== "Return") continue;
      const archiveEntry = archiveEntries[i];
      const selectedByItem = row.selectedByItem || {};
      if (row.docType === "Return") {
        for (const [itemId, codes] of Object.entries(selectedByItem)) {
          if (!codes || codes.length === 0) continue;
          const returned = new Set(codes);
          returnEntries.push({ itemId, codes: [...returned], row, archiveEntry });
        }
        continue;
      }
      const hasItemSelections = Object.values(selectedByItem).some((codes) => (codes || []).length > 0);

      if (hasItemSelections) {
        // Resolve the sheet's declared totals the same way the row displayed them, so the
        // saved delivery carries the figure the user was looking at rather than a
        // recount of the packing list.
        const deliveryLots = Object.keys(selectedByItem)
          .filter((id) => (selectedByItem[id] || []).length > 0)
          .map((id) => (items || []).find((it) => it.id === id))
          .filter(Boolean);
        const deliveryDeclaredDist = distributeDeclaredAcrossLots(
          row.declaredTotalsList,
          deliveryLots,
          (it) => sumSelectedPackages(it.packages, selectedByItem[it.id] || []),
          row.declaredByGroup || {}
        );
        for (const [itemId, codes] of Object.entries(selectedByItem)) {
          if (!codes || codes.length === 0) continue;
          const itemUnitCode = ((items || []).find((x) => x.id === itemId) || {}).unitCode;
          const fromGroup = deliveryDeclaredDist[itemId];
          const declared = (fromGroup && fromGroup.split)
            ? fromGroup
            : ((row.declaredByItem || {})[itemId] || fromGroup || null);
          existingDeliveryEntries.push({
            itemId,
            delivery: {
              date: row.date || todayStr(), deliveredTo: row.projectEn || row.projectZh, receivedBy: "",
              jobNumber: row.jobNumber, recordedBy: "", notes: t.legacyImportedNote(legacySourceName(row)),
              shkNumber: row.shkNumber || "",
              codes,
              declared: declared && (declaredNum(declared.kg) != null || declaredNum(declared.cbm) != null)
                ? { kg: declared.kg || "", cbm: declared.cbm || "", split: !!declared.split }
                : null,
              oversize: (() => {
                const detected = (row.oversizeByLot || {})[itemUnitCode];
                const os = normaliseOversize((row.oversizeByDeliveryItem || {})[itemId]
                  || (detected && detected.length ? { checked: true, cases: detected } : null));
                const list = os.checked ? cleanOversizeCases(os.cases) : [];
                return list.length ? { cases: list, cbm: String(oversizeCbmTotal(list)) } : null;
              })(),
            },
            archiveEntry,
          });
        }
        continue;
      }

      const refNos = String(row.referJobNumber || "").split(",").map((s) => s.trim()).filter(Boolean);
      if (!refNos.length) continue;
      const deliveryRecord = {
        date: row.date || todayStr(), deliveredTo: row.projectEn || row.projectZh, receivedBy: "",
        jobNumber: row.jobNumber, recordedBy: "", notes: t.legacyImportedNote(legacySourceName(row)),
        shkNumber: row.shkNumber || "",
        packageCount: row.packageCount || 1,
      };
      let anyRefMatched = false;
      for (const refNo of refNos) {
        const sameBatchIdx = jobNoToImportIdx.get(refNo);
        if (sameBatchIdx != null) {
          importRows[sameBatchIdx].deliveries = [...(importRows[sameBatchIdx].deliveries || []), { ...deliveryRecord, id: `D${Date.now()}${i}` }];
          archiveEntry.__linkedItemIndex = sameBatchIdx;
          sameBatchDeliveryCount++;
          anyRefMatched = true;
          continue;
        }
        const existing = items.find((it) => String(it.jobNumber || "").trim() === refNo);
        if (existing) {
          existingDeliveryEntries.push({ itemId: existing.id, delivery: deliveryRecord, archiveEntry });
          anyRefMatched = true;
        } else {
          archiveEntry.unmatchedReferral = refNo;
        }
      }
    }

    const createdItems = importRows.length > 0 ? onLegacyImport(importRows) : [];
    for (const entry of archiveEntries) {
      if (entry.__linkedItemIndex != null) entry.linkedItemId = createdItems[entry.__linkedItemIndex] ? createdItems[entry.__linkedItemIndex].id : null;
      delete entry.__linkedItemIndex;
      delete entry.rowIndex;
    }
    if (pendingEnrichments.length > 0 && onLegacyEnrich) onLegacyEnrich(pendingEnrichments);
    if (existingDeliveryEntries.length > 0 && onLegacyDeliver) {
      const results = onLegacyDeliver(existingDeliveryEntries.map((e) => ({ itemId: e.itemId, delivery: e.delivery })));
      results.forEach((r, idx) => {
        const entry = existingDeliveryEntries[idx].archiveEntry;
        if (!entry.linkedItemIds.includes(r.itemId)) entry.linkedItemIds.push(r.itemId);
        entry.linkedItemId = entry.linkedItemIds.join(", ");
      });
    }

    // Returns, applied last so they act on entries the deliveries in this same batch have
    // already touched. A returned case is taken back off whichever delivery took it out -
    // the goods came back, so that case was not delivered after all - and the entry holds it
    // again. A delivery left with no cases at all is cancelled rather than kept as an empty
    // record.
    if (returnEntries.length > 0 && onLegacyEnrich) {
      const patches = [];
      for (const r of returnEntries) {
        const it = (items || []).find((x) => x.id === r.itemId);
        if (!it) continue;
        const back = new Set(r.codes.map((c) => String(c).trim()));
        let touched = 0;
        const deliveries = (it.deliveries || []).map((d) => {
          const keep = (d.codes || []).filter((c) => !back.has(String(c).trim()));
          if (keep.length === (d.codes || []).length) return d;
          touched += (d.codes || []).length - keep.length;
          // The record stays either way, even when every case came back. The truck still went
          // to site and back, and that trip is charged at the minimum whatever returned with
          // it - cancelling the delivery would erase a movement that really happened and a
          // charge that is really owed. What changes is the cases it carried.
          return {
            ...d, codes: keep,
            allReturned: keep.length === 0,
            notes: [d.notes, t.legacyReturnedNote(r.row.jobNumber)].filter(Boolean).join(" \u00b7 "),
          };
        });
        if (!touched) continue;
        patches.push({ itemId: r.itemId, patch: { deliveries } });
        if (!r.archiveEntry.linkedItemIds.includes(r.itemId)) r.archiveEntry.linkedItemIds.push(r.itemId);
        r.archiveEntry.linkedItemId = r.archiveEntry.linkedItemIds.join(", ");
      }
      if (patches.length) onLegacyEnrich(patches);
    }

    setLegacyArchive((prev) => [...archiveEntries, ...prev]);
    setResults({ archived: archiveEntries.length, created: createdItems.length, delivered: sameBatchDeliveryCount + existingDeliveryEntries.length, enriched: pendingEnrichments.length, checkedIn: checkInOps.length });
    setRows([]);
    setProcessing(false);
  }

  // Each dropdown offers only what has actually been uploaded, so a filter can never come
  // back empty. Job numbers run in order - they are issued sequentially, so reading them
  // in order is how a gap or a duplicate shows itself.
  const backlogOptions = (() => {
    const grab = (key) => [...new Set((legacyArchive || []).map((r) => String(r[key] || "").trim()).filter(Boolean))];
    return {
      jobNumbers: grab("jobNumber").sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
      sites: grab("project").sort((a, b) => a.localeCompare(b)),
      jobRefs: grab("jobRef").sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
      clients: grab("client").sort((a, b) => a.localeCompare(b)),
      fileNames: grab("fileName").sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
    };
  })();
  const backlogFiltered = (legacyArchive || []).filter((r) => {
    if (backlogTypeFilter !== "All" && r.docType !== backlogTypeFilter) return false;
    if (backlogJobFilter !== "All" && String(r.jobNumber || "").trim() !== backlogJobFilter) return false;
    if (backlogSiteFilter !== "All" && String(r.project || "").trim() !== backlogSiteFilter) return false;
    if (backlogJobRefFilter !== "All" && String(r.jobRef || "").trim() !== backlogJobRefFilter) return false;
    if (backlogClientFilter !== "All" && String(r.client || "").trim() !== backlogClientFilter) return false;
    if (backlogFileFilter !== "All" && String(r.fileName || "").trim() !== backlogFileFilter) return false;
    if (!backlogSearch.trim()) return true;
    const q = backlogSearch.toLowerCase();
    return [r.fileName, r.client, r.project, r.jobNumber, r.jobRef, r.docType, r.linkedItemId, r.unmatchedReferral]
      .some((v) => String(v || "").toLowerCase().includes(q));
  });
  // Job numbers are issued in sequence, so reading the list in that order is how a gap or a
  // duplicate shows itself. Uploads arrive in whatever order the files were found, which is
  // why that is not the default but is a click away.
  const backlogSorted = backlogSort === "jobNumber"
    ? [...backlogFiltered].sort((a, b) =>
        String(a.jobNumber || "\uffff").localeCompare(String(b.jobNumber || "\uffff"), undefined, { numeric: true })
        || String(a.fileName || "").localeCompare(String(b.fileName || ""), undefined, { numeric: true }))
    : backlogFiltered;
  const backlogFiltersOn = [backlogTypeFilter, backlogJobFilter, backlogSiteFilter, backlogJobRefFilter,
    backlogClientFilter, backlogFileFilter].some((v) => v !== "All") || !!backlogSearch.trim();
  function clearBacklogFilters() {
    setBacklogSearch("");
    setBacklogTypeFilter("All");
    setBacklogJobFilter("All");
    setBacklogSiteFilter("All");
    setBacklogJobRefFilter("All");
    setBacklogClientFilter("All");
    setBacklogFileFilter("All");
  }

  // Drops the listing and the copy of the file kept with it. Storage has no delete, so the
  // key is emptied the same way an invoice document is.
  function removeArchiveRow(id) {
    storageSet(`legacyDoc:${id}`, "").catch(() => {});
    setLegacyArchive((prev) => prev.filter((row) => row.id !== id));
    setDeletingBacklogId(null);
    setDeletePlan(null);
  }
  async function viewArchivedFile(id) {
    try {
      const res = await storageGet(`legacyDoc:${id}`);
      if (!res) return;
      const { uri, name } = JSON.parse(res.value);
      const w = window.open("", "_blank");
      if (!w) return;
      if (uri.startsWith("data:application/pdf")) {
        w.document.write(`<title>${name}</title><embed src="${uri}" type="application/pdf" style="width:100%;height:100vh;">`);
      } else {
        w.document.write(`<title>${name}</title><body style="margin:0;background:#333;display:flex;justify-content:center;"><img src="${uri}" style="max-width:100%;height:auto;"></body>`);
      }
    } catch (e) { /* noop */ }
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.legacyUploadTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.legacyUploadDesc}</p>
        <input ref={fileInputRef} type="file" multiple accept=".xlsx,.xls,.pdf,image/*" className="hidden" onChange={handleFilesSelected} disabled={scanning} />
        <button
          className="px-3 py-1.5 rounded text-sm font-semibold"
          style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: scanning ? 0.6 : 1 }}
          disabled={scanning}
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
        >
          {scanning ? t.legacyScanningMsg : t.legacyChooseFilesBtn}
        </button>
        {!scanning && <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>{t.legacyAutoDetectHint}</div>}
      </div>

      {rows.length > 0 && (
        <div className="flex flex-col gap-3">
          {rows.map((row, idx) => (
            <LegacyUploadRow key={idx} onReplaceIncomingCases={onReplaceIncomingCases} directory={directory} setDirectory={setDirectory} employees={employees} row={row} siblingRows={rows} onChange={(next) => updateRow(idx, next)} onRemove={() => removeRow(idx)} incoming={incoming} items={items} onLegacyEnrich={onLegacyEnrich} onAddIncoming={onAddIncoming} onProcessAll={processAll} processing={processing} processDisabled={processing || rows.some(legacyRowMissing)} colors={colors} t={t} lang={lang} />
          ))}
          {rows.some((r) => !r.client) && (
            <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
              {t.legacyClientRequiredSummaryMsg}
            </div>
          )}
          {rows.some((r) => !r.projectEn && !r.projectZh) && (
            <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
              {t.legacySiteRequiredSummaryMsg}
            </div>
          )}
          <div className="flex gap-2">
            <button
              className="px-4 py-2 rounded text-sm font-semibold"
              style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: (processing || rows.some((r) => !r.projectEn && !r.projectZh) || rows.some((r) => !r.client)) ? 0.6 : 1 }}
              disabled={processing || rows.some((r) => !r.projectEn && !r.projectZh) || rows.some((r) => !r.client)}
              onClick={processAll}
            >
              {processing ? t.legacyProcessingMsg : t.legacyProcessBtn(rows.length)}
            </button>
          </div>
        </div>
      )}
      {results && (
        <div className="px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>
          {t.legacyResultsMsg(results.archived, results.created, results.delivered, results.enriched, results.checkedIn)}
        </div>
      )}

      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.legacyBacklogTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.legacyBacklogDesc}</p>
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.searchLabel} colors={colors}>
            <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }} value={backlogSearch} onChange={(e) => setBacklogSearch(e.target.value)} />
          </Field>
          <Field label={t.fJobNumber} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={backlogJobFilter} onChange={(e) => setBacklogJobFilter(e.target.value)}>
              <option value="All">{t.statusAll}</option>
              {backlogOptions.jobNumbers.map((v) => <option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label={t.legacyProjectSite} colors={colors}>
            <select className={inputClass} style={{ ...inputStyleFor(colors), maxWidth: 240 }} value={backlogSiteFilter} onChange={(e) => setBacklogSiteFilter(e.target.value)}>
              <option value="All">{t.statusAll}</option>
              {backlogOptions.sites.map((v) => <option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label={t.fJobRef} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={backlogJobRefFilter} onChange={(e) => setBacklogJobRefFilter(e.target.value)}>
              <option value="All">{t.statusAll}</option>
              {backlogOptions.jobRefs.map((v) => <option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label={t.clientLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={backlogClientFilter} onChange={(e) => setBacklogClientFilter(e.target.value)}>
              <option value="All">{t.statusAll}</option>
              {backlogOptions.clients.map((v) => <option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label={t.legacyDocType} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={backlogTypeFilter} onChange={(e) => setBacklogTypeFilter(e.target.value)}>
              <option value="All">{t.statusAll}</option>
              {LEGACY_DOC_TYPES.map((tp) => <option key={tp}>{tp}</option>)}
            </select>
          </Field>
          <Field label={t.legacyColFile} colors={colors}>
            <select className={inputClass} style={{ ...inputStyleFor(colors), maxWidth: 240 }} value={backlogFileFilter} onChange={(e) => setBacklogFileFilter(e.target.value)}>
              <option value="All">{t.statusAll}</option>
              {backlogOptions.fileNames.map((v) => <option key={v}>{v}</option>)}
            </select>
          </Field>
          <Field label={t.legacyBacklogSortLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={backlogSort} onChange={(e) => setBacklogSort(e.target.value)}>
              <option value="recent">{t.legacyBacklogSortRecent}</option>
              <option value="jobNumber">{t.legacyBacklogSortJobNo}</option>
            </select>
          </Field>
          {backlogFiltersOn && (
            <button className="text-xs font-semibold pb-2" style={{ color: colors.amberText }} onClick={clearBacklogFilters}>
              {t.clearBtn}
            </button>
          )}
        </div>
        <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>
          {t.legacyBacklogCount(backlogFiltered.length, (legacyArchive || []).length)}
        </div>
      </div>

      {pickedIds.length > 0 && (
        <div className="rounded-lg px-4 py-2 mb-2 flex flex-wrap items-center gap-3" style={{ background: colors.navy, color: colors.onDark }}>
          <span className="text-sm font-semibold" style={{ fontFamily: FONT_DISPLAY }}>{t.legacySelectedCount(pickedIds.length)}</span>
          <button className="text-sm font-semibold" style={{ color: colors.amber }} onClick={() => deletePicked(true)}>
            {t.legacyBulkReverseBtn}
          </button>
          <button className="text-sm font-semibold" style={{ color: "#CFC9BB" }} onClick={() => deletePicked(false)}>
            {t.legacyBulkDeleteBtn}
          </button>
          <button className="text-sm ml-auto" style={{ color: "#9AA0AE" }} onClick={() => setPicked({})}>{t.legacyClearSelection}</button>
        </div>
      )}
      <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
        <table className="w-full text-sm" style={{ background: colors.surface }}>
          <thead>
            <tr style={{ background: colors.surfaceDim }}>
              <th className="px-3 py-2" style={{ width: 34 }}>
                <input type="checkbox"
                  title={t.legacySelectAllHint}
                  checked={backlogSorted.length > 0 && backlogSorted.every((r) => picked[r.id])}
                  onChange={(e) => {
                    const next = { ...picked };
                    backlogSorted.forEach((r) => { if (e.target.checked) next[r.id] = true; else delete next[r.id]; });
                    setPicked(next);
                  }} />
              </th>
              {[t.legacyColFile, t.legacyDocType, t.clientLabel, t.legacyProjectSite, t.fJobNumber, t.fJobRef, t.colDate, t.legacyColLinked, ""].map((h) => (
                <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {backlogFiltered.length === 0 && (
              <tr><td colSpan={9} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{backlogFiltersOn ? t.legacyBacklogNoMatchMsg : t.legacyBacklogNoneMsg}</td></tr>
            )}
            {backlogSorted.map((r) => {
              const isEditing = editingBacklogId === r.id;
              const inputStyle = inputStyleFor(colors);
              if (deletingBacklogId === r.id && deletePlan) {
                const p = deletePlan;
                const nothingToUndo = p.items.length === 0;
                return (
                  <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, background: colors.redSoft }}>
                    <td colSpan={9} className="px-3 py-3">
                      <div className="text-xs font-semibold mb-2 truncate" style={{ color: colors.ink }}>{r.fileName}</div>
                      <div className="text-sm mb-2" style={{ color: colors.ink }}>{t.legacyDeleteTitle}</div>
                      {nothingToUndo ? (
                        <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>{t.legacyDeleteNothingMsg}</div>
                      ) : (
                        <ul className="text-xs mb-2" style={{ color: colors.ink, listStyle: "disc", paddingLeft: 18 }}>
                          {p.items.map((x) => (
                            <li key={x.itemId} className="mb-0.5">
                              {x.deliveryIds.length > 0 && t.legacyDeleteDeliveryLine(x.label, x.deliveryIds.length, x.caseCount)}
                              {x.arrivalIds.length > 0 && t.legacyDeleteArrivalLine(x.label, x.arrivalIds.length, x.caseCount)}
                              {x.remove
                                ? ` ${t.legacyDeleteEntryRemoved(x.label)}`
                                : ` ${t.legacyDeleteEntryKept(x.arrivalsLeft, x.deliveriesLeft)}`}
                            </li>
                          ))}
                          {p.incoming.map((inc) => (
                            <li key={inc.incomingId} className="mb-0.5">{t.legacyDeleteIncomingLine(inc.incomingId, inc.codes.length)}</li>
                          ))}
                        </ul>
                      )}
                      {p.stranded.length > 0 && (
                        <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>{t.legacyDeleteStrandedMsg(p.stranded.join(", "))}</div>
                      )}
                      {(p.blocked || []).map((b) => (
                        <div key={b.label} className="text-xs mb-2" style={{ color: colors.red }}>{t.legacyDeleteBlockedMsg(b.label, b.deliveries)}</div>
                      ))}
                      <div className="text-[11px] mb-3" style={{ color: colors.inkFaint }}>{t.legacyDeleteHint}</div>
                      <div>
                        {!nothingToUndo && (
                          <button
                            className="text-xs font-semibold mr-3"
                            style={{ color: colors.red }}
                            onClick={() => {
                              if (onLegacyReverse) onLegacyReverse(p);
                              removeArchiveRow(r.id);
                            }}
                          >
                            {t.legacyDeleteReverseBtn}
                          </button>
                        )}
                        <button
                          className="text-xs font-semibold mr-3"
                          style={{ color: colors.inkFaint }}
                          onClick={() => removeArchiveRow(r.id)}
                        >
                          {nothingToUndo ? t.legacyDeleteListingOnlyBtn : t.legacyDeleteKeepRecordsBtn}
                        </button>
                        <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => { setDeletingBacklogId(null); setDeletePlan(null); }}>
                          {t.cancelBtn}
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              }
              if (isEditing) {
                const d = backlogEditDraft;
                return (
                  <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, background: colors.amberSoft }}>
                    <td colSpan={9} className="px-3 py-3">
                      <div className="text-xs font-semibold mb-2 truncate" style={{ color: colors.inkFaint }}>{r.fileName}</div>
                      <div className="flex flex-wrap gap-3">
                        <Field label={t.legacyDocType} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={d.docType} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, docType: e.target.value }))}>
                            {LEGACY_DOC_TYPES.map((tp) => <option key={tp}>{tp}</option>)}
                          </select>
                        </Field>
                        <Field label={t.clientLabel} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={d.client} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, client: e.target.value }))}>
                            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                          </select>
                        </Field>
                        <Field label={t.legacyProjectSite} colors={colors}>
                          <input className={inputClass} style={inputStyle} value={d.project} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, project: e.target.value }))} />
                        </Field>
                        <Field label={t.fJobNumber} colors={colors}>
                          <input className={inputClass} style={inputStyle} value={d.jobNumber} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, jobNumber: e.target.value }))} />
                        </Field>
                        <Field label={t.colDate} colors={colors}>
                          <input type="date" className={inputClass} style={inputStyle} value={d.date} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, date: e.target.value }))} />
                        </Field>
                      </div>
                      <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>{t.legacyEditLinkedHint}</div>
                      {(d.records || []).length > 0 && (
                        <div className="mt-3 rounded p-2" style={{ border: `1px dashed ${colors.line}` }}>
                          <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                            {t.legacyLinkedRecordsLabel}
                          </div>
                          {(d.records || []).map((rec, ri) => {
                            const patchRec = (patch) => setBacklogEditDraft((p) => ({
                              ...p, records: p.records.map((x, i) => (i === ri ? { ...x, ...patch } : x)),
                            }));
                            return (
                              <div key={`${rec.originalItemId}-${rec.recId}`} className="flex flex-wrap items-end gap-2 mb-2">
                                <Field label={t.legacyRecordEntryLabel} colors={colors}>
                                  <select
                                    className={inputClass}
                                    style={{ ...inputStyle, fontSize: 12, padding: "3px 6px", minWidth: 190, color: rec.itemId === rec.originalItemId ? colors.green : colors.amberText, fontWeight: 600 }}
                                    value={rec.itemId}
                                    onChange={(e) => patchRec({ itemId: e.target.value })}
                                  >
                                    {moveTargetsFor(d.client, [rec.originalItemId, rec.itemId]).map((it) => (
                                      <option key={it.id} value={it.id}>{itemLabel(it)}</option>
                                    ))}
                                  </select>
                                </Field>
                                <Field label={t.colDate} colors={colors}>
                                  <input type="date" className={inputClass} style={{ ...inputStyle, fontSize: 12, padding: "3px 6px" }}
                                    value={rec.date} onChange={(e) => patchRec({ date: e.target.value })} />
                                </Field>
                                {rec.kind === "arrival" ? (
                                  <Field label={t.fArrivingType} colors={colors}>
                                    <select className={inputClass} style={{ ...inputStyle, fontSize: 12, padding: "3px 6px" }}
                                      value={rec.type} onChange={(e) => patchRec({ type: e.target.value })}>
                                      {ARRIVING_TYPES.map((x) => <option key={x}>{x}</option>)}
                                    </select>
                                  </Field>
                                ) : (
                                  <Field label={t.colJobNo} colors={colors}>
                                    <input className={inputClass} style={{ ...inputStyle, width: 110, fontSize: 12, padding: "3px 6px" }}
                                      value={rec.jobNumber} onChange={(e) => patchRec({ jobNumber: e.target.value })} />
                                  </Field>
                                )}
                                <Field label={t.jsKgs} colors={colors}>
                                  <input type="number" min="0" step="0.1" className={inputClass} style={{ ...inputStyle, width: 92, fontSize: 12, padding: "3px 6px" }}
                                    value={rec.kg} onChange={(e) => patchRec({ kg: e.target.value })} />
                                </Field>
                                <Field label={t.jsCbm} colors={colors}>
                                  <input type="number" min="0" step="0.001" className={inputClass} style={{ ...inputStyle, width: 88, fontSize: 12, padding: "3px 6px" }}
                                    value={rec.cbm} onChange={(e) => patchRec({ cbm: e.target.value })} />
                                </Field>
                                {rec.itemId !== rec.originalItemId && (
                                  <div className="basis-full text-[11px]" style={{ color: colors.amberText }}>
                                    {t.legacyRecordMoveNote(
                                      itemLabel((items || []).find((i) => i.id === rec.originalItemId)) || rec.originalItemId,
                                      itemLabel((items || []).find((i) => i.id === rec.itemId)) || rec.itemId,
                                      (rec.codes || []).length,
                                      codesCarriedOver(rec.codes, rec.itemId).length
                                    )}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                          <div className="text-[11px]" style={{ color: colors.inkFaint }}>{t.legacyLinkedRecordsHint}</div>
                        </div>
                      )}
                      <div className="mt-3">
                        <button
                          className="text-xs font-semibold mr-3"
                          style={{ color: colors.green }}
                          onClick={() => {
                            const { records, ...meta } = backlogEditDraft;
                            // A record that moved takes the listing's link with it, or the
                            // archive would go on pointing at the entry it just left.
                            const moves = new Map((records || [])
                              .filter((x) => x.itemId && x.originalItemId && x.itemId !== x.originalItemId)
                              .map((x) => [x.originalItemId, x.itemId]));
                            setLegacyArchive((prev) => prev.map((row) => {
                              if (row.id !== r.id) return row;
                              const next = { ...row, ...meta };
                              if (!moves.size) return next;
                              const was = (row.linkedItemIds && row.linkedItemIds.length)
                                ? row.linkedItemIds
                                : String(row.linkedItemId || "").split(",").map((s) => s.trim()).filter(Boolean);
                              const ids = [...new Set(was.map((id) => moves.get(id) || id))];
                              return { ...next, linkedItemIds: ids, linkedItemId: ids.join(", ") };
                            }));
                            saveLinkedRecords(records);
                            setEditingBacklogId(null);
                            setBacklogEditDraft(null);
                          }}
                        >
                          {t.saveBtn}
                        </button>
                        <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => { setEditingBacklogId(null); setBacklogEditDraft(null); }}>
                          {t.cancelBtn}
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              }
              return (
               <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, background: picked[r.id] ? colors.amberSoft : undefined }}>
                  <td className="px-3 py-2">
                    <input type="checkbox" checked={!!picked[r.id]}
                      onChange={(e) => setPicked((p) => ({ ...p, [r.id]: e.target.checked }))} />
                  </td>
                <td className="px-3 py-2 max-w-[200px] truncate">{r.fileName}</td>
                <td className="px-3 py-2">{r.docType}</td>
                <td className="px-3 py-2">{r.client}</td>
                <td className="px-3 py-2 max-w-[180px] truncate">{r.project}</td>
                <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.jobNumber || "—"}</td>
                <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.jobRef || "—"}</td>
                <td className="px-3 py-2">{r.date ? fmt(r.date) : "—"}</td>
                <td className="px-3 py-2">
                  {r.linkedItemId ? (
                    <span style={{ color: colors.green, fontWeight: 600 }}>{r.docType === "Delivery" ? t.legacyDeliveredFrom(r.linkedItemId) : r.linkedItemId}</span>
                  ) : r.unmatchedReferral ? (
                    <span style={{ color: colors.red }} title={t.legacyUnmatchedHint}>{t.legacyUnmatchedReferral(r.unmatchedReferral)}</span>
                  ) : (
                    <span style={{ color: colors.inkFaint }}>{t.legacyArchivedOnly}</span>
                  )}
                </td>
                <td className="px-3 py-2 text-right whitespace-nowrap">
                  {r.hasFile && <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => viewArchivedFile(r.id)}>{t.signedDocViewBtn}</button>}
                  <button
                    className="text-xs font-semibold"
                    style={{ color: colors.inkFaint }}
                    onClick={() => {
                      setEditingBacklogId(r.id);
                      setBacklogEditDraft({
                        docType: r.docType, client: r.client, project: r.project,
                        jobNumber: r.jobNumber || "", date: r.date || "",
                        records: draftRecordsFor(r),
                      });
                    }}
                  >
                    {t.editBtn}
                  </button>
                  <button
                    className="text-xs font-semibold ml-3"
                    style={{ color: colors.red }}
                    onClick={() => {
                      setEditingBacklogId(null);
                      setBacklogEditDraft(null);
                      setDeletePlan(reversalPlanFor(r));
                      setDeletingBacklogId(r.id);
                    }}
                  >
                    {t.deleteBtn}
                  </button>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DirectoryPanel({ directory, setDirectory, employees, setEmployees, freeRules, setFreeRules, cbmRates, setCbmRates, legacyArchive, setLegacyArchive, items, incoming, onLegacyImport, onLegacyDeliver, onLegacyEnrich, onLegacyReverse, onLegacyCheckIn, onLegacyCheckInBatch, colors, t, lang }) {
  const [mode, setMode] = useState("sites");
  const [editingSite, setEditingSite] = useState(null);
  const [siteForm, setSiteForm] = useState(null);
  const [editingEmp, setEditingEmp] = useState(null);
  const [empForm, setEmpForm] = useState(null);
  const [ruleDraft, setRuleDraft] = useState({ project: "", days: "" });
  const inputStyle = inputStyleFor(colors);
  const allRoles = [...new Set([...DEFAULT_ROLES, ...employees.map((e) => e.role)])];

  function addFreeRule() {
    if (!ruleDraft.project.trim() || !(Number(ruleDraft.days) > 0)) return;
    setFreeRules((rs) => [...(rs || []), { id: `FR${Date.now()}`, project: ruleDraft.project.trim(), days: Number(ruleDraft.days) }]);
    setRuleDraft({ project: "", days: "" });
  }
  function deleteFreeRule(id) {
    setFreeRules((rs) => (rs || []).filter((r) => r.id !== id));
  }

  function setRate(client, value) {
    setCbmRates((r) => ({ ...(r || {}), [client]: value }));
  }
  function resetRate(client) {
    setCbmRates((r) => {
      const next = { ...(r || {}) };
      delete next[client];
      return next;
    });
  }

  function newSiteForm() {
    return { siteEn: "", siteZh: "", client: CLIENTS[0], jobRef: "", orderedBy: "", accountOfficer: "" };
  }
  function saveSite() {
    if (!siteForm.siteEn.trim()) return;
    if (editingSite) {
      setDirectory((d) => d.map((s) => (s.id === editingSite.id ? { ...s, ...siteForm } : s)));
    } else {
      setDirectory((d) => [...d, { ...siteForm, id: `SITE${Date.now()}` }]);
    }
    setEditingSite(null);
    setSiteForm(null);
  }
  function deleteSite(id) {
    setDirectory((d) => d.filter((s) => s.id !== id));
  }

  function newEmpForm() {
    return { name: "", role: allRoles[0] || "" };
  }
  function saveEmp() {
    if (!empForm.name.trim()) return;
    if (editingEmp) {
      setEmployees((es) => es.map((e) => (e.id === editingEmp.id ? { ...e, ...empForm } : e)));
    } else {
      setEmployees((es) => [...es, { ...empForm, id: `EMP${Date.now()}` }]);
    }
    setEditingEmp(null);
    setEmpForm(null);
  }
  function deleteEmp(id) {
    setEmployees((es) => es.filter((e) => e.id !== id));
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["sites", t.tabSitesAccounts], ["employees", t.tabEmployees], ["freedays", t.tabFreeStorage], ["pricing", t.tabPricing]].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>

      {mode === "freedays" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.freeStorageTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.freeStorageDesc}</p>
            <div className="flex flex-wrap items-end gap-3">
              <Field label={t.fFreeProject} hint={t.fFreeProjectHint} colors={colors}>
                <input className={inputClass} style={{ ...inputStyle, minWidth: "260px" }} value={ruleDraft.project} onChange={(e) => setRuleDraft((d) => ({ ...d, project: e.target.value }))} />
              </Field>
              <Field label={t.fFreeDays} colors={colors}>
                <input type="number" min="1" className={inputClass} style={{ ...inputStyle, width: "90px" }} value={ruleDraft.days} onChange={(e) => setRuleDraft((d) => ({ ...d, days: e.target.value }))} />
              </Field>
              <button
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: !ruleDraft.project.trim() || !(Number(ruleDraft.days) > 0) ? 0.5 : 1 }}
                onClick={addFreeRule}
              >
                {t.freeStorageAddBtn}
              </button>
            </div>
          </div>

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.freeColProject, t.freeColDays, ""].map((h, idx) => (
                    <th key={idx} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(freeRules || []).length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.freeStorageNoneMsg}</td></tr>
                )}
                {(freeRules || []).map((r) => (
                  <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                    <td className="px-3 py-2">{r.project}</td>
                    <td className="px-3 py-2">{r.days}</td>
                    <td className="px-3 py-2 text-right">
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => deleteFreeRule(r.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {mode === "pricing" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.pricingTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.pricingDesc}</p>
          </div>

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.pricingColClient, t.pricingColRate, ""].map((h, idx) => (
                    <th key={idx} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.keys(DEFAULT_CBM_RATES).map((client) => {
                  const override = (cbmRates || {})[client];
                  const hasOverride = override != null && override !== "";
                  return (
                    <tr key={client} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                      <td className="px-3 py-2 font-semibold">{client}</td>
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-2">
                          <span style={{ color: colors.inkFaint }}>$</span>
                          <input
                            type="number" min="0" step="0.01"
                            className={inputClass}
                            style={{ ...inputStyle, width: "110px" }}
                            placeholder={String(DEFAULT_CBM_RATES[client])}
                            value={hasOverride ? override : ""}
                            onChange={(e) => setRate(client, e.target.value)}
                          />
                          <span style={{ color: colors.inkFaint }}>{t.pricingPerCbm}</span>
                        </div>
                      </td>
                      <td className="px-3 py-2 text-right">
                        {hasOverride && (
                          <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => resetRate(client)}>{t.pricingResetBtn}</button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {mode === "sites" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.dirTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.dirDesc}</p>
            <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => { setEditingSite(null); setSiteForm(newSiteForm()); }}>
              {t.dirAddBtn}
            </button>
          </div>

          {siteForm && (
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                <Field label={t.fSiteEn} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.siteEn} onChange={(e) => setSiteForm((f) => ({ ...f, siteEn: e.target.value }))} />
                </Field>
                <Field label={t.fSiteZh} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.siteZh} onChange={(e) => setSiteForm((f) => ({ ...f, siteZh: e.target.value }))} />
                </Field>
                <Field label={t.fDirClient} colors={colors}>
                  <select className={inputClass} style={inputStyle} value={siteForm.client} onChange={(e) => setSiteForm((f) => ({ ...f, client: e.target.value }))}>
                    {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                  </select>
                </Field>
                <Field label={t.fDirJobRef} hint={t.fJobRefHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.jobRef} onChange={(e) => setSiteForm((f) => ({ ...f, jobRef: e.target.value }))} />
                </Field>
                <Field label={t.fDirOrderedBy} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.orderedBy} onChange={(e) => setSiteForm((f) => ({ ...f, orderedBy: e.target.value }))} />
                </Field>
                <Field label={t.fDirOfficer} colors={colors}>
                  <select className={inputClass} style={inputStyle} value={siteForm.accountOfficer} onChange={(e) => setSiteForm((f) => ({ ...f, accountOfficer: e.target.value }))}>
                    <option value=""></option>
                    {employees.map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
                  </select>
                </Field>
              </div>
              <div className="flex gap-2 mt-4">
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={saveSite}>
                  {t.saveBtn}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setEditingSite(null); setSiteForm(null); }}>
                  {t.cancelBtn}
                </button>
              </div>
            </div>
          )}

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.dirColSite, t.dirColClient, t.dirColJobRef, t.dirColOfficer, t.dirColOrderedBy, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {directory.length === 0 && (
                  <tr><td colSpan={6} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.dirNoneMsg}</td></tr>
                )}
                {directory.map((s) => (
                  <tr key={s.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                    <td className="px-3 py-2">
                      <div>{s.siteEn}</div>
                      {s.siteZh && <div className="text-xs" style={{ color: colors.inkFaint }}>{s.siteZh}</div>}
                    </td>
                    <td className="px-3 py-2">{s.client}</td>
                    <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.jobRef || "—"}</td>
                    <td className="px-3 py-2">{s.accountOfficer || "—"}</td>
                    <td className="px-3 py-2">{s.orderedBy || "—"}</td>
                    <td className="px-3 py-2 text-right whitespace-nowrap">
                      <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => { setEditingSite(s); setSiteForm({ ...s }); }}>{t.editBtn}</button>
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => deleteSite(s.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {mode === "employees" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.empTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.empDesc}</p>
            <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => { setEditingEmp(null); setEmpForm(newEmpForm()); }}>
              {t.empAddBtn}
            </button>
          </div>

          {empForm && (
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                <Field label={t.fEmpName} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={empForm.name} onChange={(e) => setEmpForm((f) => ({ ...f, name: e.target.value }))} />
                </Field>
                <Field label={t.fEmpRole} colors={colors}>
                  <input
                    className={inputClass}
                    style={inputStyle}
                    list="role-options"
                    placeholder={t.fEmpRolePlaceholder}
                    value={empForm.role}
                    onChange={(e) => setEmpForm((f) => ({ ...f, role: e.target.value }))}
                  />
                  <datalist id="role-options">
                    {allRoles.map((r) => <option key={r} value={r} />)}
                  </datalist>
                </Field>
              </div>
              <div className="flex gap-2 mt-4">
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={saveEmp}>
                  {t.saveBtn}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setEditingEmp(null); setEmpForm(null); }}>
                  {t.cancelBtn}
                </button>
              </div>
            </div>
          )}

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.empColName, t.empColRole, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {employees.length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.empNoneMsg}</td></tr>
                )}
                {employees.map((e) => (
                  <tr key={e.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                    <td className="px-3 py-2">{e.name}</td>
                    <td className="px-3 py-2">{e.role}</td>
                    <td className="px-3 py-2 text-right whitespace-nowrap">
                      <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => { setEditingEmp(e); setEmpForm({ ...e }); }}>{t.editBtn}</button>
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => deleteEmp(e.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, tone, colors }) {
  const toneColor = { grey: colors.ink, amber: colors.amber, red: colors.red, green: colors.green }[tone || "grey"];
  return (
    <div className="rounded-lg p-4 flex flex-col gap-1" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <span className="text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{label}</span>
      <span className="text-3xl font-bold" style={{ color: toneColor, fontFamily: FONT_MONO }}>{value}</span>
    </div>
  );
}

function downloadTemplate() {
  const headerRow = FIELD_DEFS.map((f) => f.label);
  const ws = XLSX.utils.aoa_to_sheet([headerRow]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Import Template");
  XLSX.writeFile(wb, "farspeed-import-template.xlsx");
}

function exportToExcel(items) {
  const inventoryRows = items.map((it) => {
    const row = {};
    FIELD_DEFS.forEach((f) => { row[f.label] = it[f.key] ?? ""; });
    row["Status"] = deriveStatus(it).replace("_", " ");
    row["Total Units"] = totalUnits(it);
    row["Delivered Units"] = deliveredUnits(it);
    row["Remaining Units"] = remainingUnits(it);
    const info = storageInfo(it);
    row["Billable Storage Days"] = info ? info.billableDays : 0;
    return row;
  });

  const deliveryRows = [];
  items.forEach((it) => {
    (it.deliveries || []).forEach((d) => {
      deliveryRows.push({
        "Item ID": it.id, Client: it.client, Project: it.project, "Unit Code": it.unitCode,
        "Delivery Date": d.date, "Quantity Delivered": d.packageCount, "Delivered To": d.deliveredTo,
        "Received By": d.receivedBy, Notes: d.notes,
      });
    });
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventoryRows), "Inventory");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deliveryRows.length ? deliveryRows : [{ "Item ID": "", "Delivery Date": "", "Quantity Delivered": "" }]), "Deliveries");
  XLSX.writeFile(wb, `farspeed-depot-export-${todayStr()}.xlsx`);
}

function UploadPanel({ onReplaceIncomingCases, onImportRows, onAddIncoming, existingItems, directory, setDirectory, employees, legacyArchive, setLegacyArchive, items, incoming, onLegacyImport, onLegacyCheckIn, onLegacyCheckInBatch, onLegacyDeliver, onLegacyEnrich, onLegacyReverse, colors, t, lang }) {
  const [mode, setMode] = useState("packinglist");
  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["packinglist", t.uploadModePackingList], ["legacy", t.uploadModeLegacy]].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>
      {mode === "packinglist" && (
        <ImportPanel onImportRows={onImportRows} onAddIncoming={onAddIncoming} existingItems={existingItems} directory={directory} setDirectory={setDirectory} employees={employees} colors={colors} t={t} lang={lang} hideExcelMode />
      )}
      {mode === "legacy" && (
        <LegacyUploadsPanel onReplaceIncomingCases={onReplaceIncomingCases} employees={employees} setDirectory={setDirectory} legacyArchive={legacyArchive} setLegacyArchive={setLegacyArchive} items={items} incoming={incoming} onLegacyCheckIn={onLegacyCheckIn} onLegacyCheckInBatch={onLegacyCheckInBatch} directory={directory} onLegacyImport={onLegacyImport} onLegacyDeliver={onLegacyDeliver} onLegacyEnrich={onLegacyEnrich} onLegacyReverse={onLegacyReverse} onAddIncoming={onAddIncoming} colors={colors} t={t} lang={lang} />
      )}
    </div>
  );
}

// Editing a lot's weight or volume here has to reach the cases, not just the row: the
// depot stores and bills per case, and the group total is only their sum. A figure typed
// in is spread across the lot's cases in the proportion they already hold - so correcting
// a total that was read wrongly keeps the heavier cases heavier - and evenly where the
// cases carry nothing yet, as when a document gives no volume at all. The last case takes
// the rounding so the cases still add up to what was typed.
// Renumbering a lot's cases from the preview. Case numbers are the depot's handle on a
// box, and a scan or an Excel column does not always give them in the form Farspeed files
// them under - so they are editable before anything reaches Incoming.
//
// Weights and descriptions stay with their case by position, so renaming is free. Adding
// or removing a case is a real change to what the lot holds, so the lot's totals are
// recomputed from the cases that remain rather than being quietly kept.
// A shipping document often gives one volume for the whole consignment - the Chevalier
// delivery order states 24 CBM across four containers - without saying how it divides
// between the orders on it. Weight is the only thing the packing list does give per lot,
// and for guide rails and lift parts the two track each other closely enough to be a
// defensible basis. So the total is shared out in proportion to each lot's weight, and
// then across that lot's cases, with the last lot and last case absorbing the rounding so
// the parts still add to the figure on the paper.
//
// It is an apportionment, not a measurement. Where the real per-lot volumes are known they
// should be typed in instead.
function distributeCbmByWeight(groups, totalCbm) {
  const total = Number(totalCbm);
  const list = groups || [];
  if (!list.length || !isFinite(total) || total <= 0) return list;
  const weights = list.map((g) => (g.packages || []).reduce((s, p) => s + (Number(p.weightKg) || 0), 0));
  const sum = weights.reduce((a, b) => a + b, 0);
  let run = 0;
  return list.map((g, i) => {
    const last = i === list.length - 1;
    const share = sum > 0 ? weights[i] / sum : 1 / list.length;
    const cbm = last ? Math.round((total - run) * 1000) / 1000 : Math.round(total * share * 1000) / 1000;
    run += cbm;
    return spreadGroupTotal(g, "cbm", String(cbm));
  });
}
function renumberGroupCases(group, text) {
  const codes = String(text || "").split(/[,\n]/).map((c) => c.trim()).filter(Boolean);
  const old = group.packages || [];
  const next = codes.map((code, i) => (old[i]
    ? { ...old[i], code }
    : { code, description: (old[0] && old[0].description) || "", weightKg: "", cbm: "" }));
  return {
    ...group,
    packages: next,
    totalWeight: next.reduce((s, p) => s + (Number(p.weightKg) || 0), 0),
    totalCbm: next.reduce((s, p) => s + (Number(p.cbm) || 0), 0),
  };
}
// A packing list that carries two jobs ends with a Shipping Marks block for each - "FUJITEC
// / ZDZ1703 / PO NO.HE-6717 / C/NO. 02-05, 09-10, 13-30, 34-35" - which is the shipper's own
// statement of which case numbers belong to which job. Expanded here into one code per case,
// keeping the printed width so "02-05" gives 02, 03, 04, 05 rather than 2, 3, 4, 5.
function expandCaseMarkRanges(text) {
  const out = [];
  const cleaned = String(text || "").replace(/C\s*\/?\s*NO\.?|CASE\s*NO\.?|\u4ef6\u865f/gi, " ");
  for (const part of cleaned.split(/[,;\u3001\n]/)) {
    const piece = part.trim();
    if (!piece) continue;
    const range = piece.match(/^(\d+)\s*[-\u2013\u2014~]\s*(\d+)$/);
    if (range) {
      const from = Number(range[1]);
      const to = Number(range[2]);
      const width = Math.max(range[1].length, range[2].length);
      // A run longer than the whole shipment is a misread, not a range; left alone.
      if (from > 0 && to >= from && to - from < 500) {
        for (let n = from; n <= to; n += 1) out.push(String(n).padStart(width, "0"));
        continue;
      }
    }
    out.push(piece);
  }
  return out;
}
// Two jobs on one factory packing list are numbered separately and both start at 02, so they
// share most of their case numbers and differ only in the last few - one job ending 32-33 and
// the other 34-35. Reading those off the page means tracking which section each row sits in
// over nineteen pages, and the tail of the second job is exactly where that slips. The
// Shipping Marks say it outright, so where they and the cases read disagree, the marks win.
//
// Only a lot whose case count already matches its mark is corrected: a count that disagrees
// means a case was missed or duplicated, which renaming would paper over rather than fix, and
// that is reported separately. The correction is by position, which is sound because a job's
// cases are printed in ascending order and its mark lists them the same way.
function reconcileGroupCaseCodes(groups, shippingMarks) {
  const corrections = [];
  const marks = (shippingMarks || [])
    .map((m) => ({ lot: String((m && m.lot) || "").trim(), codes: expandCaseMarkRanges(m && m.cases) }))
    .filter((m) => m.lot && m.codes.length);
  if (!marks.length) return { groups: groups || [], corrections };
  const lotKey = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const sameCode = (c) => String(c || "").trim().toUpperCase().replace(/^0+(?=\d)/, "");
  const rank = (c) => {
    const n = parseFloat(String(c || "").replace(/[^\d.]/g, ""));
    return isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
  };
  const next = (groups || []).map((g) => {
    const mark = marks.find((x) => lotKey(x.lot) === lotKey(g.lot));
    const pkgs = g.packages || [];
    if (!mark || pkgs.length !== mark.codes.length || !pkgs.length) return g;
    // A mark listing the same number twice cannot say which case is which; left alone.
    if (new Set(mark.codes.map(sameCode)).size !== mark.codes.length) return g;
    if (pkgs.map((p) => sameCode(p.code)).join("|") === mark.codes.map(sameCode).join("|")) return g;
    const byPosition = pkgs.map((_, i) => i).sort((a, b) => rank(pkgs[a].code) - rank(pkgs[b].code) || a - b);
    const wanted = [...mark.codes].sort((a, b) => rank(a) - rank(b));
    const packages = pkgs.slice();
    const changed = [];
    byPosition.forEach((pkgIdx, i) => {
      if (sameCode(pkgs[pkgIdx].code) === sameCode(wanted[i])) return;
      packages[pkgIdx] = { ...pkgs[pkgIdx], code: wanted[i] };
      changed.push(`${pkgs[pkgIdx].code || "\u2014"} \u2192 ${wanted[i]}`);
    });
    if (!changed.length) return g;
    corrections.push({ lot: g.lot, changed });
    return { ...g, packages };
  });
  return { groups: next, corrections };
}
// A Chevalier project is named by a code at the front of its site name - EL-1876, REL-2205,
// FEEL-1330. Guide rails for several projects ship on one packing list, so the site cannot
// be a property of the file: the A4523 list carries EL-1924 and EL-1876, and the A4366 list
// carries EL-1926, EL-1909 and EL-1876. Each lot has to find its own site, or four lots for
// two different buildings land under one heading and every delivery afterwards has to be
// picked out of the wrong pile.
function extractProjectCode(text) {
  const m = String(text || "").match(/\b((?:FEEL|REL|EL)\s*-?\s*\d{3,5})\b/i);
  if (!m) return "";
  // "EL 1876", "el-1876" and "EL1876" are the same project; normalise to "EL-1876".
  return m[1].toUpperCase().replace(/^([A-Z]+)\s*-?\s*(\d+)$/, "$1-$2");
}
function findDirectorySiteByCode(directory, code, client) {
  if (!code) return null;
  const wanted = code.replace(/[^A-Z0-9]/gi, "").toUpperCase();
  const same = (site) => [site.siteEn, site.siteZh, site.jobRef]
    .some((v) => extractProjectCode(v).replace(/[^A-Z0-9]/gi, "").toUpperCase() === wanted);
  return (directory || []).find((d) => d.client === client && same(d))
    || (directory || []).find(same) || null;
}
function spreadGroupTotal(group, field, raw) {
  const total = Number(raw);
  const pkgs = group.packages || [];
  const totalField = field === "cbm" ? "totalCbm" : "totalWeight";
  if (!pkgs.length || !isFinite(total) || total < 0) {
    return { ...group, [totalField]: Number(raw) || 0 };
  }
  const dp = field === "cbm" ? 3 : 2;
  const f = Math.pow(10, dp);
  const current = pkgs.map((p) => Number(p[field]) || 0);
  const sum = current.reduce((a, b) => a + b, 0);
  let run = 0;
  const next = pkgs.map((p, i) => {
    const share = sum > 0 ? current[i] / sum : 1 / pkgs.length;
    const last = i === pkgs.length - 1;
    const v = last ? Math.round((total - run) * f) / f : Math.round(total * share * f) / f;
    run += v;
    return { ...p, [field]: v ? String(v) : "" };
  });
  return { ...group, packages: next, [totalField]: total };
}
// Edits a Directory site without leaving the screen you noticed the problem on. The site
// details - who orders for it, its job ref, its Chinese name - are usually wrong or missing
// at exactly the moment a packing list or job sheet is being imported against it, and
// walking over to the Directory tab to fix a name means losing the upload in progress.
function InlineSiteEditor({ site, setDirectory, employees, colors, t, defaultOpen, label }) {
  const [open, setOpen] = useState(!!defaultOpen);
  const [edits, setEdits] = useState(null);
  const [saved, setSaved] = useState(false);
  const inputStyle = inputStyleFor(colors);
  if (!site) return null;
  // The draft is derived, not stored: opening straight away via defaultOpen would otherwise
  // render before any state had been set, and reading the fields off null blanks the page.
  const form = { ...site, ...(edits || {}) };
  const setField = (k) => (e) => setEdits((prev) => ({ ...(prev || {}), [k]: e.target.value }));
  const start = () => { setEdits(null); setSaved(false); setOpen(true); };
  const save = () => {
    setDirectory((d) => (d || []).map((x) => (x.id === site.id ? { ...x, ...form } : x)));
    setEdits(null);
    setOpen(false);
    setSaved(true);
  };
  if (!open) {
    return (
      <div className="flex items-center gap-2 mt-1">
        <button type="button" className="text-xs font-semibold underline" style={{ color: colors.amberText }} onClick={start}>
          {label || t.dirInlineEditBtn}
        </button>
        {saved && <span className="text-xs" style={{ color: colors.green }}>{t.dirInlineSavedMsg}</span>}
      </div>
    );
  }
  return (
    <div className="rounded p-3 mt-2" style={{ border: `1px dashed ${colors.line}`, background: colors.surfaceDim }}>
      <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
        {t.dirInlineEditTitle}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
        <Field label={t.fSiteEn} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.siteEn || ""} onChange={setField("siteEn")} />
        </Field>
        <Field label={t.fSiteZh} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.siteZh || ""} onChange={setField("siteZh")} />
        </Field>
        <Field label={t.fDirClient} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.client || CLIENTS[0]} onChange={setField("client")}>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.fDirJobRef} hint={t.fJobRefHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.jobRef || ""} onChange={setField("jobRef")} />
        </Field>
        <Field label={t.fDirOrderedBy} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.orderedBy || ""} onChange={setField("orderedBy")} />
        </Field>
        <Field label={t.fDirOfficer} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.accountOfficer || ""} onChange={setField("accountOfficer")}>
            <option value=""></option>
            {(employees || []).map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
          </select>
        </Field>
      </div>
      <div className="flex gap-2 mt-3">
        <button type="button" className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={save}>{t.saveBtn}</button>
        <button type="button" className="px-3 py-1.5 rounded text-xs font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => { setEdits(null); setOpen(false); }}>{t.cancelBtn}</button>
      </div>
      <div className="text-[11px] mt-2" style={{ color: colors.inkFaint }}>{t.dirInlineEditHint}</div>
    </div>
  );
}
// These live at module level, not inside ImportPanel. They were nested there by accident,
// which put them out of reach of the Packing list reader on the settings menu - a different
// component entirely - and every file it read failed with "packingListSummaryRow is not
// defined". Two screens use them now, so neither owns them.
// One row of the packing list summary, from whichever kind of file it came out of. The
// reference is what everything downstream keys on - the DM number, the SHK, whatever the
// maker calls its lot - and the cases are the real content, so the two are checked against
// each other here and the answer travels with the row.
// The maker's name as the depot writes it, and as the packing list wrote it. A Mitsubishi
// list says 三菱电梯香港有限公司 and a TK one says (蒂升); the app knows them as Mitsubishi and
// TK Elevator. Both are kept, because the English is what everything downstream matches on
// and the Chinese is what the paperwork in front of you actually says.
const hasCJK = (v) => /[\u3400-\u9FFF]/.test(String(v || ""));
function clientNamePair(raw) {
  const en = resolveClientGuess(raw) || (hasCJK(raw) ? "" : String(raw || "").trim());
  const zh = hasCJK(raw) ? String(raw).trim()
    : Object.entries(CHINESE_CLIENT_ALIASES).find(([, name]) => name === en)?.[0] || "";
  return { en, zh };
}
// A site is written one way or the other on a packing list, rarely both. Whichever came off
// the sheet is kept in its own column, and the directory is asked for the other half - that
// is what the directory is for, and guessing a translation here would put a name in the file
// that appears nowhere else.
function siteNamePair(raw, directory) {
  const text = String(raw || "").trim();
  const norm = (v) => String(v || "").toLowerCase().replace(/[^a-z0-9\u3400-\u9FFF]/g, "");
  const key = norm(text);
  const hit = key && (directory || []).find((d) => {
    const a = norm(d.siteEn), b = norm(d.siteZh);
    return (a && (a === key || a.includes(key) || key.includes(a)))
      || (b && (b === key || b.includes(key) || key.includes(b)));
  });
  if (hit) return { en: hit.siteEn || (hasCJK(text) ? "" : text), zh: hit.siteZh || (hasCJK(text) ? text : "") };
  return hasCJK(text) ? { en: "", zh: text } : { en: text, zh: "" };
}
// The scan is told to join a packing list's case parts with the lift number in front -
// "B11 01 (#.01)" is 01B1101 - but a model asked to follow a format will occasionally hand
// back what it saw. Repairing it here as well costs nothing and means a marking never
// reaches the depot in a form nothing matches.
//
// Only the exact shape is touched: a component code, a suffix, and a lift number in
// parentheses, all on one marking. Anything already joined, or shaped any other way, is
// left alone - guessing more widely would corrupt the makers who number their cases plainly.
function repairLiftFirstMarking(code, lift) {
  const text = String(code == null ? "" : code).trim();
  const m = text.match(/^([A-Z]{1,2}\d{2})\s+([\dA-Z][\dA-Z-]*)\s*\(#\.(\d{1,2})\)$/i);
  if (m) return `${m[3]}${m[1]}${m[2]}`.toUpperCase().replace(/\s+/g, "");
  // The same thing with the marker already stripped off, leaving "B11 01".
  const n = text.match(/^([A-Z]{1,2}\d{2})\s+(\d{1,2}[\dA-Z-]*)$/i);
  if (n) return `${n[2].slice(0, 2)}${n[1]}${n[2]}`.toUpperCase().replace(/\s+/g, "");
  // Joined, but with the lift left off the front: "E21 23" came back as "E2123" rather than
  // "23E2123". The lift is known from the lot the case sits in, so it is put where it
  // belongs - but only when the marking does not already start with it, or a correct
  // "23B1123" would be given a second 23.
  const l = String(lift == null ? "" : lift).trim();
  if (/^\d{1,2}$/.test(l) && /^[A-Z]{1,2}\d/i.test(text) && !text.startsWith(l)) {
    return `${l}${text}`.toUpperCase().replace(/\s+/g, "");
  }
  return text.toUpperCase() === text ? text : text;
}
function packingListSummaryRow(fileName, client, project, ref, packages, kg, cbm, statedPkgs, directory) {
  // The lot doubles as the lift number on these packing lists - "23", "24" - which is what
  // a marking that came back without its prefix needs.
  const liftHint = /^\d{1,2}$/.test(String(ref || "").trim()) ? String(ref).trim() : "";
  const codes = (packages || []).map((p) => repairLiftFirstMarking((p && p.code) || "", liftHint)).filter(Boolean);
  const pkgs = statedPkgs || codes.length || (packages || []).length;
  const listed = codes.length;
  const c = clientNamePair(client);
  const site = siteNamePair(project, directory);
  return {
    "File Name": fileName,
    "Client": c.en, "Client (\u4e2d\u6587)": c.zh,
    "Project": site.en, "Project (\u4e2d\u6587)": site.zh,
    "DM or SHK or other Client Reference": ref || "",
    "PKGS": pkgs || "", "KGS": kg ? Math.round(kg * 100) / 100 : "",
    "CBM": cbm ? Math.round(cbm * 1000) / 1000 : "",
    "Cases": codes.join(", "),
    // Blank when they agree, so the eye goes straight to the ones that do not.
    "Check": listed === 0 ? "no case numbers" : (listed === pkgs ? "" : "Pkgs != Case #"),
    // The row's own verdict, kept apart from the cross-file notes. Check is rebuilt from
    // this every time; building it from the previous Check made each keystroke append
    // another copy, so typing "Mitsubishi" left one note for "M", another for "Mi", and so
    // on until the column was taller than the screen.
    __baseCheck: listed === 0 ? "no case numbers" : (listed === pkgs ? "" : "Pkgs != Case #"),
    __listed: listed,
  };
}
const PL_SUMMARY_COLUMNS = ["File Name", "Client", "Client (\u4e2d\u6587)", "Project", "Project (\u4e2d\u6587)", "DM or SHK or other Client Reference", "PKGS", "KGS", "CBM", "Cases", "Check"];
// A sheet only has to name its file, its packages and its cases to be recognised; the
// client and site columns may be in either language.
const PL_SUMMARY_REQUIRED = ["file name", "pkgs", "cases"];
function packingListSummaryColumns(headerRow) {
  const map = {};
  (headerRow || []).forEach((cell, i) => {
    const key = String(cell == null ? "" : cell).trim().toLowerCase();
    if (!key) return;
    // "DM or SHK or other Client Reference" is the reference column, however it is worded.
    // "Client (中文)" and "Project (中文)" are their own columns; the plain ones stay the keys.
    const name = /\bdm\b|shk|reference/.test(key) ? "reference"
      : /^client\s*\(/.test(key) ? "clientzh"
      : /^project\s*\(/.test(key) ? "projectzh"
      : key;
    if (map[name] === undefined) map[name] = i;
  });
  const named = map.client !== undefined || map.clientzh !== undefined;
  return PL_SUMMARY_REQUIRED.every((k) => map[k] !== undefined) && map.reference !== undefined && named ? map : null;
}
// Turns that sheet back into the lots the packing list importer already knows how to place.
// Cases are the real thing: where a lot lists them, they become its packages one for one,
// and the stated package count is only used when there are none to count.
function groupsFromPackingListSummary(grid) {
  const headerIdx = (grid || []).findIndex((r) => packingListSummaryColumns(r));
  if (headerIdx === -1) return null;
  const col = packingListSummaryColumns(grid[headerIdx]);
  const cell = (row, name) => {
    const i = col[name];
    return i === undefined ? "" : String((row || [])[i] == null ? "" : (row || [])[i]).trim();
  };
  const groups = [];
  let client = "", project = "";
  for (let i = headerIdx + 1; i < grid.length; i++) {
    const r = grid[i] || [];
    if (r.every((c) => String(c == null ? "" : c).trim() === "")) continue;
    const codes = cell(r, "cases").split(/[,&\u3001]/).map((c) => c.trim()).filter(Boolean);
    const stated = Number(String(cell(r, "pkgs")).replace(/,/g, "")) || 0;
    const kg = Number(String(cell(r, "kgs")).replace(/,/g, "")) || 0;
    const cbm = Number(String(cell(r, "cbm")).replace(/,/g, "")) || 0;
    const count = codes.length || stated;
    if (!count) continue;
    // Either column will do on the way back in; the English one is preferred because that
    // is what the depot matches on, but a sheet that only filled in the Chinese still works.
    const rowClient = cell(r, "client") || cell(r, "clientzh");
    const rowProject = cell(r, "project") || cell(r, "projectzh");
    client = client || rowClient;
    project = project || rowProject;
    // Weight and volume are stated for the lot, so they are spread evenly across its cases.
    // Nothing in this sheet says what any single case weighs.
    const packages = (codes.length ? codes : Array.from({ length: stated }, (_, n) => `${n + 1}/${stated}`))
      .map((code) => ({
        code,
        weightKg: kg ? String(Math.round((kg / count) * 100) / 100) : "",
        cbm: cbm ? String(Math.round((cbm / count) * 1000) / 1000) : "",
        description: "",
      }));
    groups.push({
      lot: cell(r, "reference") || cell(r, "file name") || `row ${i + 1}`,
      client: rowClient, project: rowProject,
      sourceFile: cell(r, "file name"),
      packages,
      totalWeight: kg, totalCbm: cbm,
      statedPkgs: stated, listedCases: codes.length,
      containers: [],
    });
  }
  return groups.length ? { groups, client, project } : null;
}

function ImportPanel({ onImportRows, onAddIncoming, existingItems, directory, setDirectory, employees, colors, t, lang, hideExcelMode }) {
  const [showOlderSites, setShowOlderSites] = useState(false);
  const [mode, setMode] = useState("packinglist");
  const [excelPreview, setExcelPreview] = useState(null);
  const [included, setIncluded] = useState([]);
  const [excelError, setExcelError] = useState("");
  const [plPreview, setPlPreview] = useState(null);
  const [plError, setPlError] = useState("");
  const [plCommon, setPlCommon] = useState(null);
  const [plExpanded, setPlExpanded] = useState(null);
  const [plShipmentCbm, setPlShipmentCbm] = useState("");
  const [newSiteIds, setNewSiteIds] = useState([]);
  const [pdfStatus, setPdfStatus] = useState("idle"); // idle | scanning
  const [pdfError, setPdfError] = useState("");
  // Terminal dates a release notice carries. They belong to the check-in rather than the
  // packing list, so they are shown for copying across rather than silently dropped.
  const [pdfTerminalDates, setPdfTerminalDates] = useState(null);
  const [pdfWarnings, setPdfWarnings] = useState([]);
  const [pdfDocumentTotals, setPdfDocumentTotals] = useState(null);
  const inputStyle = inputStyleFor(colors);
  const siteSuggestions = useMemo(() => {
    const fromDirectory = (directory || []).map((s) => s.siteEn).filter(Boolean);
    const fromItems = (existingItems || []).map((i) => i.project).filter(Boolean);
    return [...new Set([...fromDirectory, ...fromItems])];
  }, [directory, existingItems]);

  function applyParsedResult({ groups, client, project }) {
    if (!groups || groups.length === 0) { return false; }
    const resolvedClientEarly = resolveClientGuess(client);
    // Each lot looks for its own site by the project code in its name, before anything is
    // applied to all of them.
    setPlPreview(groups.map((g) => {
      const code = extractProjectCode(g.lot);
      const site = findDirectorySiteByCode(directory, code, resolvedClientEarly);
      return { ...g, projectCode: code, directoryId: site ? site.id : "" };
    }));
    const guess = String(project || "").toLowerCase();
    const matchedSite = guess ? (directory || []).find((s) =>
      (s.siteEn && guess.includes(s.siteEn.toLowerCase())) ||
      (s.siteZh && guess.includes(s.siteZh.toLowerCase())) ||
      (s.siteEn && s.siteEn.toLowerCase().includes(guess))
    ) : null;
    const resolvedClient = resolvedClientEarly;
    setPlCommon({
      client: matchedSite ? matchedSite.client : (resolvedClient || CLIENTS[CLIENTS.length - 1]),
      project: matchedSite ? matchedSite.siteEn : (project || ""),
      directoryId: matchedSite ? matchedSite.id : "",
      jobRef: matchedSite ? matchedSite.jobRef : "",
      orderedBy: matchedSite ? matchedSite.orderedBy : "",
      shkNumber: "",
      constructionSite: matchedSite ? (matchedSite.siteZh || matchedSite.siteEn) : (project || ""),
      saveToDirectory: !matchedSite,
    });
    return true;
  }

  async function handlePdfScanFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setPdfError("");
    setPdfWarnings([]);
    setPdfDocumentTotals(null);
    setPlShipmentCbm("");
    setNewSiteIds([]);
    setPdfTerminalDates(null);
    setPlPreview(null);
    setPdfStatus("scanning");
    // Only used to explain a failure afterwards: a Fujitec list runs to nineteen pages
    // because every case is followed by its contents, and that is what takes long enough
    // to time the scan out.
    const bigPdf = file.size > 700 * 1024;
    const pageHint = Math.round(file.size / 1024);
    try {
      const base64 = await new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result.split(",")[1]);
        r.onerror = () => reject(new Error("read failed"));
        r.readAsDataURL(file);
      });
      const prompt = PDF_SCAN_PROMPT;
      const parsed = await postPdfScan({
        model: "claude-sonnet-4-6",
        max_tokens: 16000,
        messages: [{ role: "user", content: [
          { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
          { type: "text", text: prompt },
        ] }],
      });
      // Work the volume out here rather than asking the scan to do the arithmetic. The
      // Gage Street packing list prints its dimensions in three separate columns headed
      // "Length cm", "Width cm", "Height cm", and the scan simply left CBM blank for all
      // twenty-five cases. Reading the numbers off is what a scan is good at; multiplying
      // them is not, so the model now returns them as printed and the sum happens here.
      //
      // A centimetre cube is a millionth of a cubic metre and a millimetre cube a
      // billionth. Where the document doesn't name its unit, the size of the numbers
      // decides: a case measured in millimetres runs to hundreds or thousands, so a
      // largest dimension under 400 is centimetres.
      const cbmFromDims = (p) => {
        const l = Number(p.length), w = Number(p.width), h = Number(p.height);
        if (!(l > 0 && w > 0 && h > 0)) return 0;
        const unit = String(p.dimUnit || "").toLowerCase();
        const divisor = unit === "cm" ? 1e6 : unit === "mm" ? 1e9 : (Math.max(l, w, h) < 400 ? 1e6 : 1e9);
        return Math.round((l * w * h / divisor) * 10000) / 10000;
      };
      // A Delivery Memo or release notice states its totals once and lists the case
      // markings separately - "29 Package(s) / 14.088 CBM / 12,909 Kgs" over a C/S NO.
      // block. Those become one case per marking here, with the stated totals spread
      // evenly across them, because the document says nothing per case. The last case
      // carries the rounding so the cases still add up to what was stated.
      // A packing list whose "package" column is a count rather than a case number:
      // "29 | T89/B | 18,461.80" is twenty-nine cases sharing that line's weight. Each
      // line's cases carry that line's weight divided between them, and where the heavier
      // of net and gross differ the heavier is taken, since that is what the depot stores
      // and bills on. A line covering zero packages - loose fastening hardware riding
      // inside another line's cases - still has weight, so it is spread across the order's
      // cases rather than dropped.
      const expandLineGroup = (g) => {
        const lines = (g.lines || []).map((l) => ({
          n: Math.max(0, Math.round(Number(l.packages) || 0)),
          description: l.description || "",
          kg: Math.max(Number(l.netWeightKg) || 0, Number(l.grossWeightKg) || 0),
          cbm: Number(l.cbm) || 0,
        }));
        const total = lines.reduce((s2, l) => s2 + l.n, 0);
        if (!total) return [];
        const orderKg = lines.reduce((s2, l) => s2 + l.kg, 0);
        const orderCbm = lines.reduce((s2, l) => s2 + l.cbm, 0);
        const out = [];
        for (const l of lines) {
          for (let i = 0; i < l.n; i += 1) {
            out.push({
              code: "", description: l.description,
              kg: l.n ? l.kg / l.n : 0,
              cbm: l.n ? l.cbm / l.n : 0,
            });
          }
        }
        // Whatever the zero-package lines contributed, shared out over every case.
        const assignedKg = out.reduce((s2, c) => s2 + c.kg, 0);
        const assignedCbm = out.reduce((s2, c) => s2 + c.cbm, 0);
        const spreadKg = (orderKg - assignedKg) / out.length;
        const spreadCbm = (orderCbm - assignedCbm) / out.length;
        let runKg = 0, runCbm = 0;
        return out.map((c, i) => {
          const last = i === out.length - 1;
          // The last case absorbs the rounding so the cases still add to the order total.
          const kg = last ? Math.round((orderKg - runKg) * 100) / 100 : Math.round((c.kg + spreadKg) * 100) / 100;
          const cbm = last ? Math.round((orderCbm - runCbm) * 1000) / 1000 : Math.round((c.cbm + spreadCbm) * 1000) / 1000;
          runKg += kg; runCbm += cbm;
          return {
            code: `${i + 1}/${out.length}`,
            description: c.description || g.description || "ELEVATOR MATERIALS",
            weightKg: kg ? String(kg) : "",
            cbm: cbm ? String(cbm) : "",
          };
        });
      };
      const expandStatedGroup = (g) => {
        const marks = (g.caseNumbers || [])
          .flatMap((m) => String(m || "").split(","))
          .map((m) => m.trim())
          .filter(Boolean);
        const stated = Number(g.statedPackages) || 0;
        // Real markings always win, even when there are fewer of them than the document
        // says. The 13-DM-26-0523 memo declares 128 packages and its attached list holds
        // 120: topping the list up to 128 with placeholders would invent eight cases that
        // nobody can find on a pallet. The gap is reported instead. Placeholders are only
        // for a document that lists no markings at all.
        const codes = marks.length ? marks
          : (stated > 0 ? Array.from({ length: stated }, (_, i) => `${i + 1}/${stated}`) : []);
        if (!codes.length) return [];
        // The stated weight and volume are the shipment's and stay whole: they are shared
        // over the cases actually listed, so the totals still add up to what arrived even
        // though the case count is short.
        const kg = Number(g.statedWeightKg) || 0;
        const cbm = Number(g.statedCbm) || 0;
        const per = (total, i, dp) => {
          if (!total) return "";
          const f = Math.pow(10, dp);
          const each = Math.round((total / codes.length) * f) / f;
          return String(i === codes.length - 1 ? Math.round((total - each * (codes.length - 1)) * f) / f : each);
        };
        return codes.map((code, i) => ({
          code, description: g.description || "ELEVATOR PARTS",
          weightKg: per(kg, i, 2), cbm: per(cbm, i, 3),
        }));
      };
      const readGroups = (parsed.groups || []).map((g) => {
        if (!(g.packages || []).length) {
          const built = (g.lines || []).length ? expandLineGroup(g) : expandStatedGroup(g);
          if (built.length) {
            return {
              // The group letter and INS reference belong in the lot name when the scan
              // found them: a packing list split A/B/C/D sends each group to different
              // lifts, and "CED-1833" alone would not say which.
              lot: g.lot || [g.group, g.insRef].filter(Boolean).join(" ") || "UNSPECIFIED",
              containers: g.containers || [],
              totalWeight: built.reduce((s2, p) => s2 + (Number(p.weightKg) || 0), 0),
              totalCbm: built.reduce((s2, p) => s2 + (Number(p.cbm) || 0), 0),
              packages: built,
            };
          }
        }
        const packages = (g.packages || []).map((p) => {
          const stated = p.cbm !== "" && p.cbm != null ? Number(p.cbm) : null;
          const cbm = stated != null && stated > 0 ? stated : cbmFromDims(p);
          return {
            code: p.code || "",
            description: p.description || "",
            weightKg: p.weightKg !== "" && p.weightKg != null ? String(p.weightKg) : "",
            cbm: cbm > 0 ? String(cbm) : "",
          };
        });
        return {
          lot: g.lot || "UNSPECIFIED",
          containers: g.containers || [],
          totalWeight: packages.reduce((s, p) => s + (Number(p.weightKg) || 0), 0),
          totalCbm: packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0),
          packages,
        };
      });
      // Where the document says outright which case numbers belong to which job, that is
      // taken over what was read off the rows. Every change is reported rather than made
      // quietly, so it can be checked against the paper.
      const reconciled = reconcileGroupCaseCodes(readGroups, parsed.shippingMarks);
      const normalizedGroups = reconciled.groups;
      // Compose the SS/D.O. line from the parts if the scan didn't put one together - a
      // release notice gives vessel, voyage, container and B/L in four separate fields.
      const sh = parsed.shipping || {};
      const ssDoNo = parsed.ssDoNo || [
        sh.vessel ? `ex ss."${sh.vessel}"` : "",
        sh.voyage ? `V.${sh.voyage}` : "",
        sh.containerNo ? `CONTAINERS NO. ${sh.containerNo}` : "",
        sh.blNo ? `B/L ${sh.blNo}` : "",
      ].filter(Boolean).join("; ");
      setPdfTerminalDates({
        terminalArrivalDate: parsed.terminalArrivalDate || "",
        lastFreeDay: parsed.lastFreeDay || "",
      });
      // A scanned document is read off a text layer that is not always faithful, so a case
      // list that disagrees with the document's own stated package count is worth a second
      // look before any of it reaches the depot record. A marking appearing twice is not
      // itself an error - Mitsubishi's 13-DM-26-0500 genuinely carries two cases marked
      // 01C01 - so it is counted, not flagged.
      const scanWarnings = [];
      for (const c of reconciled.corrections) {
        scanWarnings.push(t.pdfCasesCorrected(c.lot, c.changed.join(", ")));
      }
      for (const g of normalizedGroups) {
        const stated = Number((parsed.groups || []).find((x) => (x.lot || "UNSPECIFIED") === g.lot)?.statedPackages) || 0;
        if (stated > 0 && stated !== g.packages.length) {
          scanWarnings.push(t.pdfCaseCountMismatch(g.lot, stated, g.packages.length));
        }
      }
      // A total the document states only once, on a companion delivery order, belongs to
      // the shipment rather than to any one order on it. Reported so it can be entered
      // where it belongs, instead of being split across orders on a guess.
      const dt = parsed.documentTotals || {};
      const readPkgs = normalizedGroups.reduce((n, g) => n + g.packages.length, 0);
      // Two groups that came back with the same name would be merged by whoever reads the
      // preview, so say so rather than letting them look like one lot.
      // Two groups of one order can carry the same annotation - CED-1831's A and B are both
      // EL-1926 - and would then merge into one lot on import. The group letter the document
      // itself uses is what tells them apart, so it is appended rather than left to chance.
      const nameCounts = new Map();
      for (const g of normalizedGroups) nameCounts.set(g.lot, (nameCounts.get(g.lot) || 0) + 1);
      const usedSuffix = new Map();
      for (const g of normalizedGroups) {
        if ((nameCounts.get(g.lot) || 0) < 2) continue;
        const n = (usedSuffix.get(g.lot) || 0) + 1;
        usedSuffix.set(g.lot, n);
        const letter = g.group || String.fromCharCode(64 + n);
        g.lot = `${g.lot} \u00b7 ${letter}`;
      }
      const lotNames = normalizedGroups.map((g) => g.lot);
      const repeatedLots = [...new Set(lotNames.filter((l, i) => lotNames.indexOf(l) !== i))];
      if (repeatedLots.length) scanWarnings.push(t.pdfRepeatedLots(repeatedLots.join(", ")));
      if (Number(dt.packages) > 0 && Number(dt.packages) !== readPkgs) {
        scanWarnings.push(t.pdfCaseCountMismatch(t.pdfWholeDocument, Number(dt.packages), readPkgs));
      }
      // The document's own total is the one figure that can be checked without knowing the
      // layout. A weight that disagrees means cases were read into the wrong lots or missed
      // altogether - a swap between two lots keeps the total right, so this is a floor, not
      // a guarantee, but it catches the coarser errors before anything is imported.
      const readKg = Math.round(normalizedGroups
        .reduce((n, g) => n + g.packages.reduce((a, p) => a + (Number(p.weightKg) || 0), 0), 0) * 100) / 100;
      const statedKg = Number(dt.weightKg) || 0;
      if (statedKg > 0 && readKg > 0 && Math.abs(statedKg - readKg) > Math.max(1, statedKg * 0.005)) {
        scanWarnings.push(t.pdfWeightMismatch(statedKg, readKg));
      }
      // The whole-shipment notice is for a document that gives one volume and no per-case
      // figures. A Fujitec list prints Volume(M3) against every case, so its lots already
      // carry real volumes that add up to the stated total, and telling someone the volume
      // "has not been divided up" invites them to replace measured figures with an estimate.
      // So the notice, and the pre-filled split box, appear only when the cases fall short of
      // the total.
      const readCbm = Math.round(normalizedGroups
        .reduce((n, g) => n + g.packages.reduce((a, p) => a + (Number(p.cbm) || 0), 0), 0) * 1000) / 1000;
      const statedCbm = Number(dt.cbm) || 0;
      const cbmAlreadyPerCase = statedCbm > 0 && readCbm > 0
        && Math.abs(statedCbm - readCbm) <= Math.max(0.05, statedCbm * 0.01);
      setPdfDocumentTotals(!cbmAlreadyPerCase && (dt.cbm || dt.weightKg) ? dt : null);
      // Pre-fill the split box with the figure the document actually gave, so it is one
      // press rather than a re-typing exercise.
      if (statedCbm > 0 && !cbmAlreadyPerCase) setPlShipmentCbm(String(dt.cbm));
      setPdfWarnings(scanWarnings);
      const ok = applyParsedResult({ groups: normalizedGroups, client: parsed.client, project: parsed.project, ssDoNo });
      if (!ok) setPdfError(t.packingListNoStructure);
      setPdfStatus("idle");
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      if (msg === "truncated-or-invalid-json" || msg === "reply-too-long") {
        setPdfError(t.pdfTruncatedMsg);
      } else if (/server returned (5\d\d|408|504)|network:|timed? ?out/i.test(msg) && bigPdf) {
        // A long packing list is the usual reason the scan never comes back.
        setPdfError(t.pdfTooLargeMsg(pageHint, msg));
      } else {
        const friendly = t.pdfReadErrorMsg || "Couldn't read this PDF. Please check the file, or enter the details manually below.";
        const detail = err && err.message ? err.message : "";
        setPdfError(detail ? `${friendly} (${detail})` : friendly);
      }
      setPdfStatus("idle");
    }
    e.target.value = "";
  }

  async function handleExcelFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setExcelError("");
    setExcelPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (rows.length === 0) { setExcelError(t.excelNoRowsMsg); return; }
      const unmatched = new Set();
      const mapped = rows.map((row) => {
        const item = emptyForm();
        Object.entries(row).forEach(([header, value]) => {
          const key = matchField(header);
          if (!key) { if (String(value).trim() !== "") unmatched.add(header); return; }
          item[key] = value instanceof Date ? dateToLocalISO(value) : String(value).trim();
        });
        return item;
      });

      const existingSigs = new Map();
      (existingItems || []).forEach((it) => { const sig = itemSignature(it); if (sig) existingSigs.set(sig, it.id); });
      const seenInBatch = new Map();
      const withDupInfo = mapped.map((item, idx) => {
        const sig = itemSignature(item);
        let dupOf = null;
        if (sig) {
          if (existingSigs.has(sig)) dupOf = existingSigs.get(sig);
          else if (seenInBatch.has(sig)) dupOf = `row ${seenInBatch.get(sig) + 1}`;
          else seenInBatch.set(sig, idx);
        }
        return { item, dupOf };
      });

      setExcelPreview({ rows: withDupInfo, unmatched: [...unmatched] });
      setIncluded(withDupInfo.map((r) => !r.dupOf));
    } catch (err) {
      setExcelError(t.excelErrorMsg);
    }
    e.target.value = "";
  }

// The columns of the packing list summary sheet: one row per lot, each naming its own
// client, site and reference. Several packing lists collapse into one file, which is then
// read back here in a single upload instead of one at a time.
  async function handlePackingListFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setPlError("");
    setPlPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      // A summary sheet carrying many packing lists is unpacked first; a single packing
      // list falls through to the reader that handles one.
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: false });
      const summary = groupsFromPackingListSummary(grid);
      const { groups, client, project } = summary || parsePackingListWorkbook(wb);
      const ok = applyParsedResult({ groups, client, project });
      if (!ok) setPlError(t.packingListNoStructure);
      // A lot whose case list is shorter or longer than the count it declares is the one
      // thing this sheet cannot fix for itself, and it is what puts the wrong number of
      // cases into the depot. Named on the way in, while it can still be corrected.
      else if (summary) {
        const off = groups.filter((g) => g.listedCases > 0 && g.statedPkgs > 0 && g.listedCases !== g.statedPkgs);
        if (off.length) {
          setPlError(t.packingListCaseCountWarn(off.map((g) => `${g.lot} (${g.statedPkgs} pkgs, ${g.listedCases} cases)`).join("; ")));
        }
      }
    } catch (err) {
      setPlError(t.packingListNoStructure);
    }
    e.target.value = "";
  }

  // Creates the Directory entry a lot is asking for, straight from this screen. The code is
  // all the packing list gives - "EL-1909" - so the site is created under that name and the
  // editor opens beneath the table to put the real building name and orderer against it. A
  // guide-rail list can bring in three projects at once, none of them on file yet.
  function addSiteForLot(idx) {
    const g = plPreview[idx];
    if (!g || !g.projectCode) return;
    const existing = findDirectorySiteByCode(directory, g.projectCode, plCommon ? plCommon.client : "");
    const site = existing || {
      id: `SITE${Date.now()}${idx}`,
      siteEn: g.projectCode,
      siteZh: "",
      client: plCommon ? plCommon.client : CLIENTS[0],
      jobRef: "",
      orderedBy: (plCommon && plCommon.orderedBy) || "",
      accountOfficer: "",
    };
    if (!existing) setDirectory((d) => [...(d || []), site]);
    // Every lot on this file with the same code gets it, not just the row that was pressed.
    setPlPreview((prev) => prev.map((row) => (
      row.projectCode === g.projectCode && !row.directoryId ? { ...row, directoryId: site.id } : row)));
    setNewSiteIds((prev) => (prev.includes(site.id) ? prev : [...prev, site.id]));
  }

  function addToIncoming() {
    let effectiveDirectoryId = plCommon.directoryId || "";

    if (plCommon.saveToDirectory && !plCommon.directoryId && plCommon.project) {
      const newSite = {
        id: `SITE${Date.now()}`,
        siteEn: plCommon.project,
        siteZh: plCommon.constructionSite && plCommon.constructionSite !== plCommon.project ? plCommon.constructionSite : "",
        client: plCommon.client,
        jobRef: plCommon.jobRef || "",
        orderedBy: plCommon.orderedBy || "",
        accountOfficer: "",
      };
      setDirectory((d) => [...d, newSite]);
      effectiveDirectoryId = newSite.id;
    }

    // A lot that resolved to its own site keeps it; the rest fall back to the common
    // fields above, which is the whole file when it only covers one project.
    const newIncoming = plPreview.map((g) => {
      const site = g.directoryId ? (directory || []).find((d) => d.id === g.directoryId) : null;
      return {
        client: plCommon.client,
        project: site ? site.siteEn : plCommon.project,
        constructionSite: site ? (site.siteZh || site.siteEn) : (plCommon.constructionSite || ""),
        jobRef: site ? (site.jobRef || "") : (plCommon.jobRef || ""),
        orderedBy: site ? (site.orderedBy || plCommon.orderedBy || "") : (plCommon.orderedBy || ""),
        shkNumber: plCommon.shkNumber || "",
        directoryId: site ? site.id : effectiveDirectoryId,
        unitCode: g.lot,
        packages: g.packages,
        notes: g.containers.length ? `Container(s): ${g.containers.join(", ")}` : "",
      };
    });
    onAddIncoming(newIncoming);
    setPlPreview(null);
    setPlCommon(null);
    setNewSiteIds([]);
  }




  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["packinglist", t.tabPackingList], ["pdf", t.tabPdf], ...(hideExcelMode ? [] : [["excel", t.tabExcel]])].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>

      {mode === "packinglist" && (
        <div className="rounded-lg p-5 flex flex-col gap-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
          <div>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.packingListTitle}</h3>
            <p className="text-sm" style={{ color: colors.inkFaint }}>{t.packingListDesc}</p>
          </div>
          <label className="px-3 py-2 rounded text-sm font-semibold cursor-pointer w-fit" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}>
            {t.choosePackingListBtn}
            <input type="file" accept=".xlsx,.xls,.xlsm,.csv" className="hidden" onChange={handlePackingListFile} />
          </label>
          {plError && <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{plError}</div>}
        </div>
      )}

      {mode === "pdf" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5 flex flex-col gap-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <div>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.pdfTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.pdfDesc}</p>
            </div>

            <label className="px-3 py-2 rounded text-sm font-semibold cursor-pointer w-fit" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}>
              {t.choosePdfBtn}
              <input type="file" accept="application/pdf" className="hidden" onChange={handlePdfScanFile} />
            </label>
            {pdfStatus === "scanning" && <div className="text-sm" style={{ color: colors.inkFaint }}>{t.scanningMsg}</div>}
            {pdfError && <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{pdfError}</div>}
            {pdfWarnings.length > 0 && (
              <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
                {pdfWarnings.map((w, i) => <div key={i}>{w}</div>)}
              </div>
            )}
            {pdfDocumentTotals && (
              <div className="px-3 py-2 rounded text-sm" style={{ background: colors.amberSoft, color: colors.amberText }}>
                {t.pdfDocumentTotals(pdfDocumentTotals.cbm || "\u2014", pdfDocumentTotals.weightKg || "\u2014")}
              </div>
            )}
            {pdfTerminalDates && (pdfTerminalDates.terminalArrivalDate || pdfTerminalDates.lastFreeDay) && (
              <div className="px-3 py-2 rounded text-sm" style={{ background: colors.amberSoft, color: colors.amberText }}>
                {t.pdfTerminalDatesFound(pdfTerminalDates.terminalArrivalDate || "\u2014", pdfTerminalDates.lastFreeDay || "\u2014")}
              </div>
            )}
          </div>
        </div>
      )}

      {plPreview && plCommon && (
            <div className="flex flex-col gap-4">
              <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
                <h4 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
                  {t.packingListCommonFieldsTitle}
                </h4>
                {(directory || []).length > 0 && (
                  <div className="mb-4">
                    <Field label={t.selectFromDirectory} colors={colors}>
                      <select
                        className={inputClass}
                        style={inputStyle}
                        value={plCommon.directoryId || ""}
                        onChange={(e) => {
                          const site = directory.find((s) => s.id === e.target.value);
                          if (!site) return;
                          setPlCommon((c) => ({
                            ...c,
                            directoryId: site.id,
                            client: CLIENTS.includes(site.client) ? site.client : c.client,
                            project: site.siteEn || c.project,
                            constructionSite: site.siteZh || site.siteEn || c.constructionSite,
                            jobRef: site.jobRef || c.jobRef,
                            orderedBy: site.orderedBy || c.orderedBy,
                            saveToDirectory: false,
                          }));
                        }}
                      >
                        <option value="">{t.selectFromDirectoryPlaceholder}</option>
                        {visibleDirectory(directory, { client: plCommon.client, showOlder: showOlderSites, items: existingItems }).map((s) => (
                          <option key={s.id} value={s.id}>{s.siteEn} — {s.client}</option>
                        ))}
                      </select>
                    </Field>
                    <label className="flex items-center gap-1.5 text-xs mt-1.5" style={{ color: colors.inkFaint }}>
                      <input type="checkbox" checked={showOlderSites} onChange={(e) => setShowOlderSites(e.target.checked)} />
                      {t.showOlderJobs}
                      {!showOlderSites && hiddenSiteCount(directory, { client: plCommon.client, items: existingItems }) > 0 && (
                        <span>{t.showOlderJobsCount(hiddenSiteCount(directory, { client: plCommon.client, items: existingItems }))}</span>
                      )}
                    </label>
                    <InlineSiteEditor
                      site={(directory || []).find((d) => d.id === plCommon.directoryId)}
                      setDirectory={setDirectory} employees={employees} colors={colors} t={t}
                    />
                  </div>
                )}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                  <Field label={t.packingListApplyClient} colors={colors}>
                    <select className={inputClass} style={inputStyle} value={plCommon.client} onChange={(e) => setPlCommon((c) => ({ ...c, client: e.target.value }))}>
                      {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label={t.packingListApplyProject} colors={colors}>
                    <input list="site-name-suggestions" className={inputClass} style={inputStyle} value={plCommon.project} onChange={(e) => setPlCommon((c) => ({ ...c, project: e.target.value, directoryId: "" }))} />
                    <datalist id="site-name-suggestions">
                      {siteSuggestions.map((s) => <option key={s} value={s} />)}
                    </datalist>
                  </Field>
                  <Field label={t.fOrderedBy} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.orderedBy || ""} onChange={(e) => setPlCommon((c) => ({ ...c, orderedBy: e.target.value }))} />
                  </Field>
                  <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.jobRef || ""} onChange={(e) => setPlCommon((c) => ({ ...c, jobRef: e.target.value }))} />
                  </Field>
                  <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.shkNumber || ""} onChange={(e) => setPlCommon((c) => ({ ...c, shkNumber: e.target.value }))} />
                  </Field>
                </div>
                <div className="mt-3 px-3 py-2 rounded text-xs" style={{ background: colors.surfaceDim, color: colors.inkFaint }}>
                  {t.packingListIncomingHint}
                </div>
                {!plCommon.directoryId && plCommon.project && (
                  <label className="flex items-center gap-2 mt-4 text-sm" style={{ color: colors.ink }}>
                    <input
                      type="checkbox"
                      checked={!!plCommon.saveToDirectory}
                      onChange={(e) => setPlCommon((c) => ({ ...c, saveToDirectory: e.target.checked }))}
                    />
                    {t.saveNewSiteToDirectory(plCommon.project)}
                  </label>
                )}
              </div>

              <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-2 text-sm font-semibold" style={{ background: colors.amberSoft, color: colors.amberText, fontFamily: FONT_DISPLAY }}>
                  {t.packingListDetectedTitle(plPreview.length)}
                </div>
                {(() => {
                  // Guide rails for several projects ship together, so a file covering more
                  // than one is normal rather than an error - but it has to be noticed.
                  const codes = [...new Set(plPreview.map((g) => g.projectCode).filter(Boolean))];
                  if (codes.length < 2) return null;
                  return (
                    <div className="px-4 py-2 text-xs" style={{ background: colors.redSoft, color: colors.red }}>
                      {t.packingListMultiProjectHint(codes.length)} {codes.join(" \u00b7 ")}
                    </div>
                  );
                })()}
                {plPreview.length > 1 && (
                  <div className="px-4 py-2 flex flex-wrap items-end gap-2" style={{ borderBottom: `1px solid ${colors.surfaceDim}` }}>
                    <Field label={t.packingListShipmentCbmLabel} colors={colors}>
                      <input
                        type="number" min="0" step="0.001"
                        className={inputClass}
                        style={{ ...inputStyleFor(colors), width: 110 }}
                        value={plShipmentCbm}
                        onChange={(e) => setPlShipmentCbm(e.target.value)}
                      />
                    </Field>
                    <button
                      type="button"
                      className="px-3 py-2 rounded text-xs font-semibold"
                      style={{
                        background: Number(plShipmentCbm) > 0 ? colors.amber : colors.line,
                        color: Number(plShipmentCbm) > 0 ? colors.ink : colors.inkFaint,
                        fontFamily: FONT_DISPLAY,
                        cursor: Number(plShipmentCbm) > 0 ? "pointer" : "not-allowed",
                      }}
                      disabled={!(Number(plShipmentCbm) > 0)}
                      onClick={() => setPlPreview((prev) => distributeCbmByWeight(prev, plShipmentCbm))}
                    >
                      {t.packingListDistributeCbmBtn}
                    </button>
                    <span className="text-xs pb-2" style={{ color: colors.inkFaint }}>{t.packingListDistributeCbmHint}</span>
                  </div>
                )}
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {[t.colLot, t.packingListColProject, t.colPackages, t.colContainers, t.colWeight, t.colCbm, ""].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {plPreview.map((g, idx) => (
                      <React.Fragment key={idx}>
                      <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <td className="px-3 py-2">
                          <input
                            className={inputClass}
                            style={{ ...inputStyleFor(colors), minWidth: 100, fontWeight: 600 }}
                            value={g.lot}
                            onChange={(e) => setPlPreview((prev) => prev.map((grp, i) => (i === idx ? { ...grp, lot: e.target.value } : grp)))}
                          />
                        </td>
                        <td className="px-3 py-2">
                          <select
                            className={inputClass}
                            style={{ ...inputStyleFor(colors), maxWidth: 210, fontSize: 12 }}
                            value={g.directoryId || ""}
                            onChange={(e) => setPlPreview((prev) => prev.map((grp, i) => (i === idx ? { ...grp, directoryId: e.target.value } : grp)))}
                          >
                            <option value="">{t.packingListProjectFromCommon}</option>
                            {visibleDirectory(directory, { client: plCommon ? plCommon.client : "", showOlder: showOlderSites, items: existingItems }).map((site) => (
                              <option key={site.id} value={site.id}>{site.siteEn}</option>
                            ))}
                          </select>
                          {!g.directoryId && g.projectCode && (
                            <div className="text-[11px] mt-0.5" style={{ color: colors.red }}>
                              {t.packingListProjectUnknown(g.projectCode)}{" "}
                              <button
                                type="button"
                                className="font-semibold underline"
                                style={{ color: colors.amberText }}
                                onClick={() => addSiteForLot(idx)}
                              >
                                {t.packingListAddSiteBtn(g.projectCode)}
                              </button>
                            </div>
                          )}
                        </td>
                        <td className="px-3 py-2">
                          <button
                            type="button"
                            className="font-semibold underline"
                            style={{ color: colors.amberText }}
                            onClick={() => setPlExpanded(plExpanded === idx ? null : idx)}
                          >
                            {g.packages.length}
                          </button>
                        </td>
                        <td className="px-3 py-2 text-xs" style={{ color: colors.inkFaint }}>{g.containers.join(", ") || "—"}</td>
                        <td className="px-3 py-2">
                          <input
                            type="number" min="0" step="0.1"
                            className={inputClass}
                            style={{ ...inputStyleFor(colors), width: 110 }}
                            value={Math.round(g.totalWeight * 100) / 100 || ""}
                            onChange={(e) => setPlPreview((prev) => prev.map((grp, i) => (i === idx ? spreadGroupTotal(grp, "weightKg", e.target.value) : grp)))}
                          />
                        </td>
                        <td className="px-3 py-2">
                          <input
                            type="number" min="0" step="0.001"
                            className={inputClass}
                            style={{ ...inputStyleFor(colors), width: 100 }}
                            value={g.totalCbm ? Math.round(g.totalCbm * 1000) / 1000 : ""}
                            placeholder="—"
                            onChange={(e) => setPlPreview((prev) => prev.map((grp, i) => (i === idx ? spreadGroupTotal(grp, "cbm", e.target.value) : grp)))}
                          />
                        </td>
                        <td className="px-3 py-2 text-right">
                          <button
                            type="button"
                            title={t.packingListRemoveGroupBtn}
                            aria-label={t.packingListRemoveGroupBtn}
                            className="w-6 h-6 rounded-full inline-flex items-center justify-center font-bold"
                            style={{ background: colors.redSoft, color: colors.red, lineHeight: 1 }}
                            onClick={() => setPlPreview((prev) => prev.filter((_, i) => i !== idx))}
                          >
                            &minus;
                          </button>
                        </td>
                      </tr>
                      {plExpanded === idx && (
                        <tr style={{ background: colors.surfaceDim }}>
                          <td colSpan={7} className="px-3 py-2">
                            <div className="text-xs mb-1" style={{ color: colors.inkFaint }}>{t.packingListCasesHint}</div>
                            <textarea
                              className={inputClass}
                              style={{ ...inputStyleFor(colors), width: "100%", minHeight: 70, fontFamily: FONT_MONO, fontSize: 12 }}
                              value={g.packages.map((p) => p.code).join(", ")}
                              onChange={(e) => setPlPreview((prev) => prev.map((grp, i) => (i === idx ? renumberGroupCases(grp, e.target.value) : grp)))}
                            />
                          </td>
                        </tr>
                      )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
                {(() => {
                  // Every Directory site this import touches, editable here. A newly created
                  // one opens straight away, since it is holding nothing but a code.
                  const ids = [...new Set(plPreview.map((g) => g.directoryId).filter(Boolean))];
                  if (!ids.length) return null;
                  return (
                    <div className="px-4 py-3" style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                      <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                        {t.packingListSitesLabel(ids.length)}
                      </div>
                      {ids.map((id) => {
                        const site = (directory || []).find((d) => d.id === id);
                        if (!site) return null;
                        return (
                          <div key={id} className="mb-1">
                            <span className="text-xs" style={{ color: colors.ink }}>{site.siteEn || site.siteZh}</span>
                            <InlineSiteEditor
                              site={site}
                              defaultOpen={newSiteIds.includes(id)}
                              setDirectory={setDirectory}
                              employees={employees}
                              colors={colors}
                              t={t}
                            />
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
              </div>

              <div className="flex gap-2">
                <button
                  className="px-4 py-2 rounded text-sm font-semibold w-fit"
                  style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: plPreview.length === 0 ? 0.5 : 1 }}
                  disabled={plPreview.length === 0}
                  onClick={addToIncoming}
                >
                  {t.packingListAddToIncomingBtn(plPreview.length)}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setPlPreview(null); setPlCommon(null); }}>
                  {t.discardBtn}
                </button>
              </div>
            </div>
      )}

      {mode === "excel" && (
        <div className="rounded-lg p-5 flex flex-col gap-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
          <div>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.excelTitle}</h3>
            <p className="text-sm" style={{ color: colors.inkFaint }}>{t.excelDesc}</p>
          </div>
          <div className="flex gap-3 items-center flex-wrap">
            <label className="px-3 py-2 rounded text-sm font-semibold cursor-pointer" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}>
              {t.chooseFileBtn}
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelFile} />
            </label>
            <button className="text-sm font-semibold underline" style={{ color: colors.amberText }} onClick={downloadTemplate}>
              {t.downloadTemplateBtn}
            </button>
          </div>
          {excelError && <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{excelError}</div>}
          {excelPreview && (
            <div className="flex flex-col gap-3">
              <div className="text-sm flex items-center gap-3 flex-wrap" style={{ color: colors.inkFaint }}>
                <span>{t.selectedCount(included.filter(Boolean).length, excelPreview.rows.length)}</span>
                <button className="underline text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => setIncluded(excelPreview.rows.map(() => true))}>{t.selectAllBtn}</button>
                <button className="underline text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => setIncluded(excelPreview.rows.map((r) => !r.dupOf))}>{t.selectNonDupBtn}</button>
              </div>
              {excelPreview.unmatched.length > 0 && (
                <div className="px-3 py-2 rounded text-sm" style={{ background: colors.amberSoft, color: colors.amberText }}>
                  {t.unmatchedMsg}{excelPreview.unmatched.join(", ")}
                </div>
              )}
              <div className="rounded overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
                <table className="w-full text-xs" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {["", t.prevColClient, t.prevColProject, t.prevColItemType, t.prevColDepot, t.prevColDepotArrival, t.prevColMatch].map((h) => (
                        <th key={h} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {excelPreview.rows.map((r, idx) => (
                      <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, opacity: included[idx] ? 1 : 0.5 }}>
                        <td className="px-2 py-1.5">
                          <input type="checkbox" checked={!!included[idx]} onChange={(e) => setIncluded((prev) => prev.map((v, i) => (i === idx ? e.target.checked : v)))} />
                        </td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.client}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.project}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.itemType}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.depot}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.depotArrivalDate}</td>
                        <td className="px-2 py-1.5">{r.dupOf ? <Badge tone="amber" colors={colors}>{t.badgeDupOf(r.dupOf)}</Badge> : <Badge tone="green" colors={colors}>{t.badgeNew}</Badge>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-2">
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
                  onClick={() => {
                    const rowsToImport = excelPreview.rows.filter((_, idx) => included[idx]).map((r) => r.item);
                    onImportRows(rowsToImport);
                    setExcelPreview(null);
                    setIncluded([]);
                  }}>
                  {t.importBtn(included.filter(Boolean).length)}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setExcelPreview(null); setIncluded([]); }}>
                  {t.discardBtn}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

async function ensureUserAccountsSeeded() {
  try {
    const res = await storageGet("userAccounts");
    if (res) return JSON.parse(res.value);
  } catch (e) { /* not seeded yet */ }
  const defaultHash = await hashPassword(DEFAULT_PASSWORD);
  const seeded = LOGIN_ACCOUNTS.map((name) => ({ name, passwordHash: defaultHash }));
  try { await storageSet("userAccounts", JSON.stringify(seeded)); } catch (e) {}
  return seeded;
}

function LoginScreen({ onLoggedIn, colors, t, lang }) {
  const [accounts, setAccounts] = useState(null);
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let on = true;
    ensureUserAccountsSeeded().then((accts) => { if (on) setAccounts(accts); });
    return () => { on = false; };
  }, []);

  async function handleLogin(e) {
    e.preventDefault();
    if (!name || !password) { setError(t.loginErrorMissing); return; }
    setBusy(true);
    setError("");
    try {
      const accts = accounts || (await ensureUserAccountsSeeded());
      const acct = accts.find((a) => a.name === name);
      const hash = await hashPassword(password);
      if (!acct || acct.passwordHash !== hash) {
        setError(t.loginErrorWrong);
        setBusy(false);
        return;
      }
      try {
        window.localStorage.setItem("farspeed_session", JSON.stringify({ name, at: Date.now() }));
      } catch (e2) {}
      onLoggedIn(name);
    } catch (e3) {
      setError(t.loginErrorWrong);
    }
    setBusy(false);
  }

  const inputStyle = inputStyleFor(colors);
  return (
    <div className="fixed inset-0 flex items-center justify-center p-6" style={{ background: colors.bg }}>
      <form onSubmit={handleLogin} className="w-full max-w-sm rounded-lg p-6" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <div className="text-center mb-5">
          <div className="text-xl font-bold" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>FARSPEED CONTRACTORS LTD</div>
          <div className="text-xs mt-1" style={{ color: colors.amberText, fontFamily: FONT_DISPLAY }}>{t.appSubtitle}</div>
        </div>
        <Field label={t.loginNameLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} required>
            <option value="">{t.loginNamePlaceholder}</option>
            {LOGIN_ACCOUNTS.map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
        </Field>
        <div className="mt-3">
          <Field label={t.loginPasswordLabel} colors={colors}>
            <input type="password" className={inputClass} style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" required />
          </Field>
        </div>
        {error && (
          <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>
        )}
        <button
          type="submit"
          className="w-full mt-4 px-4 py-2.5 rounded text-sm font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: busy ? 0.6 : 1 }}
          disabled={busy}
        >
          {busy ? t.loginBusyMsg : t.loginBtn}
        </button>
        <div className="text-xs mt-4 text-center" style={{ color: colors.inkFaint }}>{t.loginDefaultPwHint}</div>
      </form>
    </div>
  );
}

// Simple "admin privilege" gate: re-enter the current user's own password before a
// destructive action is allowed to proceed. Not a separate admin role for now - just a
// friction check, per current requirements (may become a real permission later).
function AdminConfirmModal({ authUser, onConfirm, onClose, colors, t }) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const inputStyle = inputStyleFor(colors);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      const accts = await ensureUserAccountsSeeded();
      const acct = accts.find((a) => a.name === authUser);
      const hash = await hashPassword(password);
      if (!acct || acct.passwordHash !== hash) {
        setError(t.loginErrorWrong);
        setBusy(false);
        return;
      }
      onConfirm();
    } catch (e2) {
      setError(t.loginErrorWrong);
      setBusy(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-6" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="w-full max-w-sm rounded-lg p-6" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.adminConfirmTitle}</h3>
        <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>{t.adminConfirmDesc(authUser)}</p>
        <form onSubmit={handleSubmit}>
          <Field label={t.loginPasswordLabel} colors={colors}>
            <input type="password" autoFocus className={inputClass} style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} required />
          </Field>
          {error && <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>}
          <div className="flex gap-2 mt-4">
            <button type="submit" className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.red, color: "#fff", fontFamily: FONT_DISPLAY, opacity: busy ? 0.6 : 1 }} disabled={busy}>
              {busy ? t.loginBusyMsg : t.adminConfirmBtn}
            </button>
            <button type="button" className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={onClose}>
              {t.closeBtn}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function ChangePasswordModal({ name, onClose, colors, t }) {
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);
  const [busy, setBusy] = useState(false);
  const inputStyle = inputStyleFor(colors);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    if (next.length < 4) { setError(t.pwTooShortMsg); return; }
    if (next !== confirm) { setError(t.pwMismatchMsg); return; }
    setBusy(true);
    try {
      const accts = await ensureUserAccountsSeeded();
      const idx = accts.findIndex((a) => a.name === name);
      const currentHash = await hashPassword(current);
      if (idx === -1 || accts[idx].passwordHash !== currentHash) {
        setError(t.pwCurrentWrongMsg);
        setBusy(false);
        return;
      }
      const newHash = await hashPassword(next);
      const updated = accts.map((a, i) => (i === idx ? { ...a, passwordHash: newHash } : a));
      await storageSet("userAccounts", JSON.stringify(updated));
      setSuccess(true);
    } catch (e2) {
      setError(t.pwCurrentWrongMsg);
    }
    setBusy(false);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-6" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="w-full max-w-sm rounded-lg p-6" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-4" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.changePasswordTitle}</h3>
        {success ? (
          <>
            <div className="px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>{t.pwChangedMsg}</div>
            <button className="w-full mt-4 px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>{t.closeBtn}</button>
          </>
        ) : (
          <form onSubmit={handleSubmit}>
            <Field label={t.pwCurrentLabel} colors={colors}>
              <input type="password" className={inputClass} style={inputStyle} value={current} onChange={(e) => setCurrent(e.target.value)} required />
            </Field>
            <div className="mt-3">
              <Field label={t.pwNewLabel} colors={colors}>
                <input type="password" className={inputClass} style={inputStyle} value={next} onChange={(e) => setNext(e.target.value)} required />
              </Field>
            </div>
            <div className="mt-3">
              <Field label={t.pwConfirmLabel} colors={colors}>
                <input type="password" className={inputClass} style={inputStyle} value={confirm} onChange={(e) => setConfirm(e.target.value)} required />
              </Field>
            </div>
            {error && <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>}
            <div className="flex gap-2 mt-4">
              <button type="submit" className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: busy ? 0.6 : 1 }} disabled={busy}>
                {busy ? t.loginBusyMsg : t.pwSaveBtn}
              </button>
              <button type="button" className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={onClose}>
                {t.closeBtn}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}

export default function FarspeedInventory() {
  const [authUser, setAuthUser] = useState(undefined); // undefined = checking session, null = logged out, string = logged in
  const [changePasswordOpen, setChangePasswordOpen] = useState(false);
  const [items, setItems] = useState([]);
  const [invoices, setInvoicesState] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("dashboard");
  const [editing, setEditing] = useState(null);
  const [exitingItems, setExitingItems] = useState([]);
  const [deliveryPickerSelection, setDeliveryPickerSelection] = useState([]);
  const [deliverySearch, setDeliverySearch] = useState("");
  const [deliveryFilterClient, setDeliveryFilterClient] = useState("All");
  const [deliveryFilterDepot, setDeliveryFilterDepot] = useState("All");
  const [deliveryFilterStatus, setDeliveryFilterStatus] = useState("All");
  const [printJobSheet, setPrintJobSheet] = useState(null);
  const [expandedRowId, setExpandedRowId] = useState(null);
  // Clearing a site out of inventory is the first step of a rebuild, and one row at a time
  // is a hundred confirmations. Selection follows the filters, so ticking the header box
  // takes the rows on screen and never one a filter is hiding.
  const [pickedRows, setPickedRows] = useState({});
  const [openSite, setOpenSite] = useState({});
  const [siteTotalsOpen, setSiteTotalsOpen] = useState(true);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [newEntryMenuOpen, setNewEntryMenuOpen] = useState(false);
  const [directory, setDirectoryState] = useState([]);
  const [employees, setEmployeesState] = useState([]);
  const [freeRules, setFreeRulesState] = useState([]);
  const [cbmRates, setCbmRatesState] = useState({});
  const [legacyArchive, setLegacyArchiveState] = useState([]);
  const [incoming, setIncomingState] = useState([]);
  // The revision each shared record set was last read or written at. A save carries it
  // back so the database can refuse the write if someone else has saved since.
  const revs = useRef({});
  const [conflictKey, setConflictKey] = useState("");
  const [currentUser, setCurrentUserState] = useState("");
  const [filterClient, setFilterClient] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterDepot, setFilterDepot] = useState("All");
  const [search, setSearch] = useState("");
  const [error, setError] = useState("");
  const [lang, setLang] = useState("en");
  const [theme, setTheme] = useState("light");

  const t = TEXT[lang];
  const colors = theme === "dark" ? DARK_COLORS : LIGHT_COLORS;

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem("farspeed_session");
      if (raw) {
        const session = JSON.parse(raw);
        if (session && session.name && LOGIN_ACCOUNTS.includes(session.name)) {
          setAuthUser(session.name);
          return;
        }
      }
    } catch (e) {}
    setAuthUser(null);
  }, []);

  function handleLoggedIn(name) {
    setAuthUser(name);
    setCurrentUser(name);
  }
  function handleLogout() {
    try { window.localStorage.removeItem("farspeed_session"); } catch (e) {}
    setAuthUser(null);
  }

  useEffect(() => {
    (async () => {
      try {
        const res = await storageGet("items");
        revs.current.items = res ? res.rev : null;
        setItems(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setItems([]);
      }
      try {
        const res = await storageGet("directory");
        setDirectoryState(res ? JSON.parse(res.value) : DEFAULT_DIRECTORY);
        if (!res) storageSet("directory", JSON.stringify(DEFAULT_DIRECTORY));
      } catch (e) {
        setDirectoryState(DEFAULT_DIRECTORY);
      }
      try {
        const res = await storageGet("employees");
        setEmployeesState(res ? JSON.parse(res.value) : DEFAULT_EMPLOYEES);
        if (!res) storageSet("employees", JSON.stringify(DEFAULT_EMPLOYEES));
      } catch (e) {
        setEmployeesState(DEFAULT_EMPLOYEES);
      }
      try {
        const res = await storageGet("freeStorageRules");
        const rules = res ? JSON.parse(res.value) : [];
        setFreeRulesState(rules);
        setFreeStorageRulesGlobal(rules);
      } catch (e) {
        setFreeRulesState([]);
        setFreeStorageRulesGlobal([]);
      }
      try {
        const res = await storageGet("cbmRates", true);
        const rates = res ? JSON.parse(res.value) : {};
        setCbmRatesState(rates);
        setCbmRateOverridesGlobal(rates);
      } catch (e) {
        setCbmRatesState({});
        setCbmRateOverridesGlobal({});
      }
      try {
        const res = await storageGet("legacyArchive");
        setLegacyArchiveState(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setLegacyArchiveState([]);
      }
      try {
        const res = await storageGet("invoices");
        setInvoicesState(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setInvoicesState([]);
      }
      try {
        const res = await storageGet("incoming");
        revs.current.incoming = res ? res.rev : null;
        setIncomingState(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setIncomingState([]);
      }
      try {
        setCurrentUserState(window.localStorage.getItem("farspeed_current_user") || "");
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  function setDirectory(updater) {
    setDirectoryState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("directory", JSON.stringify(next));
      return next;
    });
  }
  function setInvoices(updater) {
    setInvoicesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("invoices", JSON.stringify(next));
      return next;
    });
  }
  function setEmployees(updater) {
    setEmployeesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("employees", JSON.stringify(next));
      return next;
    });
  }
  function setFreeRules(updater) {
    setFreeRulesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("freeStorageRules", JSON.stringify(next));
      setFreeStorageRulesGlobal(next);
      return next;
    });
  }
  function setCbmRates(updater) {
    setCbmRatesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("cbmRates", JSON.stringify(next), true);
      setCbmRateOverridesGlobal(next);
      return next;
    });
  }
  function setLegacyArchive(updater) {
    setLegacyArchiveState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("legacyArchive", JSON.stringify(next));
      return next;
    });
  }
  // Replaces a shipment's case numbers with the ones a job sheet states. The job sheets
  // are typed; a scanned Delivery Memo is read off a text layer that drops digits - the
  // 13-DM-26-0500 scan turned "01C2101" into "01C01" and "02C2102" into "02C02", leaving
  // the shipment holding 01C01 twice. Only the codes change: descriptions, weights and
  // volumes stay on their cases, matched by position.
  function handleReplaceIncomingCases(incomingId, codes) {
    const list = (codes || []).map((c) => String(c || "").trim()).filter(Boolean);
    if (!list.length) return;
    setIncoming((prev) => prev.map((inc) => {
      if (inc.id !== incomingId) return inc;
      const old = inc.packages || [];
      if (list.length !== old.length) return inc;
      // First occurrence wins: the very duplicate being corrected would otherwise map the
      // old code to whatever the *last* of its copies became, so a case already checked in
      // as 01C01 would be renamed to 01C2101 rather than left alone.
      const renamed = new Map();
      old.forEach((p, i) => { if (!renamed.has(p.code)) renamed.set(p.code, list[i]); });
      return {
        ...inc,
        packages: old.map((p, i) => ({ ...p, code: list[i] })),
        // Anything already checked in was checked in under the old numbering.
        checkedInCodes: (inc.checkedInCodes || []).map((c) => renamed.get(c) || c),
      };
    }));
  }

  function setIncoming(updater) {
    const next = typeof updater === "function" ? updater(incoming) : updater;
    persistGuarded("incoming", next, setIncomingState, (v) => JSON.parse(v));
  }
  // Adds shipments from a Packing List upload to the Incoming pool - informational only,
  // not yet checked into the depot via Devan or CFS.
  function handleAddIncoming(rows) {
    let counter = incoming.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const newRows = rows.map((r) => {
      counter += 1;
      return { ...r, numericId: counter, id: `INC-${String(counter).padStart(4, "0")}`, checkedInCodes: [], linkedItemId: null, createdAt: todayStr() };
    });
    setIncoming((prev) => [...prev, ...newRows]);
    return newRows;
  }
  // Checks selected cases from an Incoming shipment into the depot via Devan or CFS.
  // First check-in for a shipment creates the real inventory item (carrying the full case
  // pool so later partial check-ins can keep adding arrival batches to the same item,
  // reusing the existing Split Arrival mechanism); later check-ins just add another batch.
  // Batch-safe version of handleCheckIn for when several check-ins happen in one go (e.g.
  // Legacy Upload processing several Devan/CFS rows at once). Calling handleCheckIn
  // synchronously in a loop would have each call read the same stale `items`/`incoming`
  // snapshot, risking duplicate FS-#### ids or one check-in silently overwriting another's
  // arrival batch. This computes everything against one working copy instead.
  function handleCheckInBatch(operations) {
    let workingItems = [...items];
    let counter = workingItems.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const incomingPatches = new Map();
    const results = [];
    for (const op of operations) {
      const inc = incoming.find((i) => i.id === op.incomingId);
      if (!inc || op.codes.length === 0) { results.push(null); continue; }
      const batch = {
        id: `ARR${Date.now()}${Math.floor(Math.random() * 1000)}`,
        date: op.date, type: op.type, codes: op.codes,
        declared: op.declared || null,
        declaredSource: op.declaredSource || "",
      };
      const prior = incomingPatches.get(op.incomingId);
      const effectiveLinkedId = prior ? prior.linkedItemId : inc.linkedItemId;
      let resultItemId;
      if (!effectiveLinkedId) {
        counter += 1;
        const totalWeight = inc.packages.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
        const totalCbm = inc.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0);
        const opOversizeCases = cleanOversizeCases(op.oversizeCases);
        const oversizeCbmVal = op.isOversize ? (oversizeCbmTotal(opOversizeCases) || Number(op.oversizeCbm) || 0) : 0;
        const effectiveCbm = oversizeCbmVal > 0 ? oversizeCbmVal : totalCbm;
        const newItem = {
          ...emptyForm(),
          client: inc.client, project: inc.project, constructionSite: inc.constructionSite || "",
          jobRef: op.jobRef || inc.jobRef || "", orderedBy: inc.orderedBy || "", directoryId: inc.directoryId || "",
          itemType: "Separate Items", unitCode: inc.unitCode || "",
          depot: op.depot, depotArrivalDate: op.date, arrivingType: op.type, jobNumber: op.jobNumber,
          ssDoNo: op.type === "Devan" ? (op.ssDoNo || "") : "",
          shkNumber: op.shkNumber || inc.shkNumber || "",
          weightKg: totalWeight ? String(Math.round(totalWeight * 10) / 10) : "",
          volumeCbm: effectiveCbm ? String(Math.round(effectiveCbm * 1000) / 1000) : "",
          isOversize: oversizeCbmVal > 0,
          oversizeCases: opOversizeCases,
          oversizeCbm: oversizeCbmVal > 0 ? String(oversizeCbmVal) : "",
          packages: inc.packages, arrivals: [batch], deliveries: [],
          notes: t.incomingCheckedInNote(inc.id), numericId: counter,
          id: `FS-${String(counter).padStart(4, "0")}`, createdAt: todayStr(),
        };
        workingItems = [...workingItems, recomputeItemTotals(newItem)];
        resultItemId = newItem.id;
      } else {
        workingItems = workingItems.map((it) => (it.id === effectiveLinkedId ? recomputeItemTotals({ ...it, arrivals: [...(it.arrivals || []), batch] }) : it));
        resultItemId = effectiveLinkedId;
      }
      incomingPatches.set(op.incomingId, { linkedItemId: resultItemId, addedCodes: [...(prior ? prior.addedCodes : []), ...op.codes] });
      results.push(resultItemId);
    }
    if (operations.length > 0) {
      persist(workingItems);
      setIncoming((prev) => prev.map((inc) => {
        const patch = incomingPatches.get(inc.id);
        if (!patch) return inc;
        return { ...inc, checkedInCodes: [...new Set([...(inc.checkedInCodes || []), ...patch.addedCodes])], linkedItemId: patch.linkedItemId };
      }));
    }
    return results;
  }
  function handleCheckIn({ incomingId, codes, type, depot, jobNumber, date, ssDoNo, declared, declaredSource }) {
    const inc = incoming.find((i) => i.id === incomingId);
    if (!inc || codes.length === 0) return null;
    const batch = {
      id: `ARR${Date.now()}${Math.floor(Math.random() * 1000)}`,
      date, type, codes,
      declared: declared || null,
      declaredSource: declaredSource || "",
    };
    let resultItemId = inc.linkedItemId;

    if (!inc.linkedItemId) {
      let counter = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0) + 1;
      const totalWeight = inc.packages.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
      const totalCbm = inc.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0);
      const newItem = {
        ...emptyForm(),
        client: inc.client,
        project: inc.project,
        constructionSite: inc.constructionSite || "",
        jobRef: inc.jobRef || "",
        orderedBy: inc.orderedBy || "",
        directoryId: inc.directoryId || "",
        itemType: "Separate Items",
        unitCode: inc.unitCode || "",
        depot,
        depotArrivalDate: date,
        arrivingType: type,
        jobNumber,
        ssDoNo: type === "Devan" ? (ssDoNo || "") : "",
        shkNumber: inc.shkNumber || "",
        weightKg: totalWeight ? String(Math.round(totalWeight * 10) / 10) : "",
        volumeCbm: totalCbm ? String(Math.round(totalCbm * 1000) / 1000) : "",
        packages: inc.packages,
        arrivals: [batch],
        deliveries: [],
        notes: t.incomingCheckedInNote(inc.id),
        numericId: counter,
        id: `FS-${String(counter).padStart(4, "0")}`,
        createdAt: todayStr(),
      };
      persist([...items, recomputeItemTotals(newItem)]);
      resultItemId = newItem.id;
    } else {
      persist(items.map((it) => (it.id === inc.linkedItemId ? recomputeItemTotals({ ...it, arrivals: [...(it.arrivals || []), batch] }) : it)));
    }

    setIncoming((prev) => prev.map((i) => (i.id === incomingId ? { ...i, checkedInCodes: [...new Set([...(i.checkedInCodes || []), ...codes])], linkedItemId: resultItemId } : i)));
    return resultItemId;
  }
  function setCurrentUser(name) {
    setCurrentUserState(name);
    try {
      window.localStorage.setItem("farspeed_current_user", name);
    } catch (e) {}
  }

  function handleResetDeliveries() {
    if (!window.confirm(t.resetConfirmMsg)) return;
    persist(items.map((i) => ({ ...i, deliveries: [] })));
    window.alert(t.resetDoneMsg);
  }

  // Applies `next` locally, then writes it only if nobody else has saved this key since we
  // last read it. On conflict the local change is rolled back to whatever is actually on
  // the server - overwriting someone else's entries silently is the failure this exists to
  // prevent, and a lost keystroke is far cheaper than a lost depot record.
  async function persistGuarded(key, next, applyLocal, parse) {
    applyLocal(next);
    try {
      const res = await storageSetGuarded(key, JSON.stringify(next), revs.current[key]);
      if (res && res.ok) {
        revs.current[key] = res.rev;
        setError("");
        setConflictKey("");
        return true;
      }
      if (res && res.conflict) {
        const server = res.current;
        revs.current[key] = server ? server.rev : null;
        try {
          applyLocal(server ? parse(server.value) : []);
        } catch (e) {}
        setConflictKey(key);
        return false;
      }
      setError(t.saveErrorMsg);
      return false;
    } catch (e) {
      setError(t.saveErrorMsg);
      return false;
    }
  }

  async function persist(next) {
    return persistGuarded("items", next, setItems, (v) => JSON.parse(v));
  }

  function nextId() {
    const max = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    return { numericId: max + 1, id: `FS-${String(max + 1).padStart(4, "0")}` };
  }

  function handleSave(form) {
    if (editing) {
      // An edit to the arrival batches or the case list changes what the headline weight
      // and volume are a sum of, so they are recomputed. An edit that only touches other
      // fields leaves them alone, so a figure someone typed in by hand survives.
      const key = (o) => JSON.stringify(o || []);
      const recompute = key(editing.arrivals) !== key(form.arrivals)
        || key((editing.packages || []).map((p) => p.code)) !== key((form.packages || []).map((p) => p.code));
      const next = { ...editing, ...form };
      persist(items.map((i) => (i.id === editing.id ? (recompute ? recomputeItemTotals(next) : next) : i)));
    } else {
      const idFields = nextId();
      persist([...items, { ...idFields, ...form, createdAt: todayStr() }]);
    }
    setEditing(null);
    setView("inventory");
  }

  // Moves cases from one entry to another - the repair for a lot that was filed under the
  // wrong lift. Only cases still at the depot can move: one already delivered is part of a
  // delivery record, and quietly pulling it out of that record would rewrite history.
  // The case takes its place in the destination's earliest arrival batch, or the
  // destination would treat it as not yet landed and refuse to let it leave again.
  function handleMoveCases({ fromId, toId, codes }) {
    const wanted = new Set(codes || []);
    const from = items.find((i) => i.id === fromId);
    const to = items.find((i) => i.id === toId);
    if (!from || !to || from.id === to.id || wanted.size === 0) return;
    const movable = remainingPackages(from).filter((p) => wanted.has(p.code));
    if (!movable.length) return;
    const moving = new Set(movable.map((p) => p.code));
    const denom = (c) => { const m = String(c).match(/\/(\d+)\s*$/); return m ? Number(m[1]) : 0; };
    persist(items.map((i) => {
      if (i.id === fromId) {
        return recomputeItemTotals({
          ...i,
          packages: (i.packages || []).filter((p) => !moving.has(p.code)),
          arrivals: (i.arrivals || []).map((a) => ({ ...a, codes: (a.codes || []).filter((c) => !moving.has(c)) })),
        });
      }
      if (i.id === toId) {
        const already = new Set((i.packages || []).map((p) => p.code));
        const add = movable.filter((p) => !already.has(p.code));
        const next = {
          ...i,
          packages: [...(i.packages || []), ...add]
            .sort((a, b) => denom(a.code) - denom(b.code) || codeLeadingNumber(a.code) - codeLeadingNumber(b.code)),
        };
        if (usesArrivalBatches(i)) {
          const addCodes = add.map((p) => p.code);
          next.arrivals = (i.arrivals || []).map((a, idx) =>
            idx === 0 ? { ...a, codes: [...new Set([...(a.codes || []), ...addCodes])] } : a);
        }
        return recomputeItemTotals(next);
      }
      return i;
    }));
  }

  function handleDelete(id) {
    persist(items.filter((i) => i.id !== id));
  }
  // Deleting several at once has to be one pass. Calling handleDelete in a loop computed
  // every removal from the same snapshot of `items`, so each write undid the one before it
  // and only the last entry actually went.
  function handleDeleteMany(ids) {
    const gone = new Set(ids || []);
    if (!gone.size) return;
    persist(items.filter((i) => !gone.has(i.id)));
  }

  function handleAddDelivery(delivery, itemId) {
    const record = { ...delivery, id: `D${Date.now()}${Math.floor(Math.random() * 1000)}` };
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: [...(i.deliveries || []), record] } : i)));
    setExitingItems((prev) => prev.map((i) => (i.id === itemId ? { ...i, deliveries: [...(i.deliveries || []), record] } : i)));
    return record;
  }

  // Records one combined delivery across multiple items at once (same job number, date,
  // destination and receiver) - each item gets its own delivery entry with its own codes,
  // linked only by sharing that job number/date, and prints as a single combined job sheet.
  // Merges Legacy Upload metadata (SS/D.O., notes, missing site names) onto an item
  // that already exists - used when the job number matches an item created earlier via
  // the proper Packing List import, so its real per-case package data is never touched
  // or duplicated, only enriched.
  function handleLegacyEnrich(entries) {
    const byItemId = new Map(entries.map((e) => [e.itemId, e.patch]));
    // A patch that only fills in metadata (SS/D.O., site names) leaves the figures alone -
    // recomputing there would overwrite a weight someone had corrected by hand. A patch
    // that changes which cases the entry holds must recompute, because the packing-list
    // weight and volume are sums over exactly those cases.
    persist(items.map((i) => {
      if (!byItemId.has(i.id)) return i;
      const patch = byItemId.get(i.id);
      const next = { ...i, ...patch };
      return (patch.packages || patch.arrivals) ? recomputeItemTotals(next) : next;
    }));
    return entries.map((e) => ({ itemId: e.itemId }));
  }

  // Undoes what one legacy file put into the depot. An arrival is lifted off its entry, so
  // the entry falls back to what the packing list says and stops counting as arrived; the
  // cases it checked in are handed back to the Incoming shipment they came from, which is
  // where they were before the file was processed. A delivery is lifted off, and the cases
  // it took out return to store. An entry that exists only because of this file, with
  // nothing else arrived on it and nothing delivered off it, goes with the file.
  //
  // Removal is real rather than a cancelled flag: a cancelled delivery still shows in the
  // entry's history, and the point here is to put things back as though the file had never
  // been uploaded.
  function handleLegacyReverse(plan) {
    if (!plan) return { items: 0, removed: 0 };
    const byId = new Map((plan.items || []).map((p) => [p.itemId, p]));
    const removeIds = new Set((plan.items || []).filter((p) => p.remove).map((p) => p.itemId));
    const next = items
      .filter((i) => !removeIds.has(i.id))
      .map((i) => {
        const p = byId.get(i.id);
        if (!p) return i;
        const arrivals = (i.arrivals || []).filter((a) => !(p.arrivalIds || []).includes(a.id));
        const deliveries = (i.deliveries || []).filter((d) => !(p.deliveryIds || []).includes(d.id));
        const dates = arrivals.map((a) => a.date).filter(Boolean).sort();
        return recomputeItemTotals({ ...i, arrivals, deliveries, depotArrivalDate: dates[0] || i.depotArrivalDate });
      });
    persist(next);
    setExitingItems((prev) => prev.filter((i) => !removeIds.has(i.id)));
    if ((plan.incoming || []).length) {
      setIncoming((prev) => prev.map((inc) => {
        const p = (plan.incoming || []).find((x) => x.incomingId === inc.id);
        if (!p) return inc;
        const drop = new Set(p.codes || []);
        const nextInc = { ...inc, checkedInCodes: (inc.checkedInCodes || []).filter((c) => !drop.has(c)) };
        if (p.unlink) nextInc.linkedItemId = null;
        return nextInc;
      }));
    }
    return { items: (plan.items || []).length, removed: removeIds.size };
  }

  // Undoes a single check-in from the audit screen. The plan is the same shape the legacy
  // backlog builds, so the same handler does the work: the batch comes off the entry, its
  // cases go back to the Incoming shipment they came from, and an entry left holding
  // nothing goes with it.
  // Reversing a duplicate that has been delivered from would leave those deliveries with no
  // stock behind them. So the delivery records are carried across to the entry being kept
  // first, and only then is the duplicate undone - one action, in the right order, instead
  // of a note on a piece of paper and a hope that it gets re-entered.
  function mergeCheckInInto(item, arrival, keepId) {
    const keep = (items || []).find((i) => i.id === keepId);
    const moving = activeDeliveries(item);
    if (keep && moving.length) {
      const have = new Set((keep.packages || []).map((p) => String(p.code || "").trim()));
      const carried = moving.map((d) => ({
        ...d,
        // Cases the kept entry does not hold would point at nothing, so only the ones it has
        // travel; the record's own date, job number and stated figures come across whole.
        codes: (d.codes || []).filter((c) => have.has(String(c).trim())),
        notes: [d.notes, t.checkInsMergedNote(item.id)].filter(Boolean).join(" \u00b7 "),
      }));
      handleLegacyEnrich([{ itemId: keep.id, patch: { deliveries: [...(keep.deliveries || []), ...carried] } }]);
    }
    reverseOneCheckIn({ ...item, deliveries: [] }, arrival);
  }
  function reverseOneCheckIn(item, arrival) {
    // Older batches predate arrival ids, so identity falls back to the object itself rather
    // than to an id that is undefined on every one of them and would match them all.
    const picked = arrival ? [arrival] : (item.arrivals || []);
    const isPicked = (a) => picked.includes(a) || (a.id && picked.some((x) => x.id === a.id));
    const arrivalIds = picked.map((a) => a.id).filter(Boolean);
    const remainingArrivals = (item.arrivals || []).filter((a) => !isPicked(a));
    const codes = picked.flatMap((a) => a.codes || []);
    // The entry goes when its last check-in does. Keeping it because it still had deliveries
    // against it did nothing useful: an entry's cases live on its packages list, not on its
    // arrival batches, so taking the batch away left all 44 cases and 30 remaining exactly
    // where they were and the screen did not change. Whoever pressed the button has already
    // been told the deliveries go with it.
    const remove = remainingArrivals.length === 0;
    const plan = {
      items: [{ itemId: item.id, arrivalIds, deliveryIds: [], remove }],
      incoming: (incoming || [])
        .filter((inc) => inc.linkedItemId === item.id)
        .map((inc) => ({ incomingId: inc.id, codes, unlink: remove })),
    };
    handleLegacyReverse(plan);
  }

  function handleAddCombinedDelivery(entries) {
    const records = entries.map(({ itemId, delivery }) => ({ itemId, record: { ...delivery, id: `D${Date.now()}${Math.floor(Math.random() * 10000)}-${itemId}` } }));
    const byItemId = new Map();
    for (const r of records) {
      if (!byItemId.has(r.itemId)) byItemId.set(r.itemId, []);
      byItemId.get(r.itemId).push(r.record);
    }
    persist(items.map((i) => (byItemId.has(i.id) ? { ...i, deliveries: [...(i.deliveries || []), ...byItemId.get(i.id)] } : i)));
    setExitingItems((prev) => prev.map((i) => (byItemId.has(i.id) ? { ...i, deliveries: [...(i.deliveries || []), ...byItemId.get(i.id)] } : i)));
    return records;
  }

  // A recorded delivery could previously only be cancelled outright. Now it can be
  // corrected in place - date, destination, receiver, job number, the cases that went, and
  // the weight and volume the job sheet stated.
  function handleUpdateDelivery(deliveryId, itemId, patch) {
    const apply = (list) => list.map((i) => (i.id === itemId
      ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, ...patch } : d)) }
      : i));
    persist(apply(items));
    setExitingItems((prev) => apply(prev));
  }

  function handleDeleteDelivery(deliveryId, itemId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: true } : d)) } : i)));
    setExitingItems((prev) => prev.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: true } : d)) } : i)));
  }

  function handleCancelItem(itemId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, cancelled: true } : i)));
  }

  function handleRestoreItem(itemId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, cancelled: false } : i)));
  }

  function handleRestoreDelivery(itemId, deliveryId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: false } : d)) } : i)));
  }

  function handlePermanentlyDeleteDelivery(itemId, deliveryId) {
    if (!window.confirm(t.permanentDeleteConfirmMsg)) return;
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).filter((d) => d.id !== deliveryId) } : i)));
  }

  function handlePermanentlyDeleteItem(itemId) {
    if (!window.confirm(t.permanentDeleteConfirmMsg)) return;
    persist(items.filter((i) => i.id !== itemId));
  }
  // Deleting several at once is not the single delete run in a loop. That asks once per
  // entry - eleven dialogs to dismiss - and worse, every call filters the same unchanged
  // list and saves it, so the last one written wins and ten of the eleven come back. One
  // confirmation, asked by the caller, and one save.
  function handlePermanentlyDeleteItems(itemIds) {
    const gone = new Set(itemIds || []);
    if (!gone.size) return;
    persist(items.filter((i) => !gone.has(i.id)));
  }

  function handleImportRows(rows) {
    let counter = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const newItems = rows.map((r) => {
      counter += 1;
      return { ...r, numericId: counter, id: `FS-${String(counter).padStart(4, "0")}`, createdAt: todayStr() };
    });
    persist([...items, ...newItems]);
    setView("inventory");
  }

  // Same as handleImportRows but stays on the current screen and returns the created
  // items (with their assigned FS ids) - used by the Legacy Uploads bulk importer.
  function handleLegacyImport(rows) {
    let counter = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const newItems = rows.map((r) => {
      counter += 1;
      return { ...r, numericId: counter, id: `FS-${String(counter).padStart(4, "0")}`, createdAt: todayStr() };
    });
    persist([...items, ...newItems]);
    return newItems;
  }

  function handleKeepOne(groupIds, keepId) {
    persist(items.filter((i) => !(groupIds.includes(i.id) && i.id !== keepId)));
  }

  function handleDeleteGroup(groupIds) {
    persist(items.filter((i) => !groupIds.includes(i.id)));
  }

  const activeItemsList = useMemo(() => items.filter((i) => !i.cancelled), [items]);

  const filtered = useMemo(() => {
    return activeItemsList
      .filter((i) => filterClient === "All" || i.client === filterClient)
      .filter((i) => filterStatus === "All" || deriveStatus(i) === filterStatus)
      .filter((i) => filterDepot === "All" || i.depot === filterDepot)
      .filter((i) => {
        if (!search.trim()) return true;
        const q = search.toLowerCase();
        return (
          i.project?.toLowerCase().includes(q) || i.description?.toLowerCase().includes(q) ||
          i.client?.toLowerCase().includes(q) || i.id?.toLowerCase().includes(q) ||
          i.shkNumber?.toLowerCase().includes(q) || i.unitCode?.toLowerCase().includes(q) ||
          i.constructionSite?.toLowerCase().includes(q) || i.invoiceNumber?.toLowerCase().includes(q) ||
          i.jobNumber?.toLowerCase().includes(q) || i.ssDoNo?.toLowerCase().includes(q) ||
          (i.packages || []).some((p) => (p.code || "").toLowerCase().includes(q)) ||
          (i.deliveries || []).some((d) => (d.jobNumber || "").toLowerCase().includes(q))
        );
      })
      .sort((a, b) => b.numericId - a.numericId);
  }, [activeItemsList, filterClient, filterStatus, filterDepot, search]);

  const atDepot = activeItemsList.filter((i) => deriveStatus(i) === "at_depot");
  const partial = activeItemsList.filter((i) => deriveStatus(i) === "partial");
  const openForDelivery = [...atDepot, ...partial];
  const filteredForDelivery = openForDelivery.filter((i) => {
    if (deliveryFilterClient !== "All" && i.client !== deliveryFilterClient) return false;
    if (deliveryFilterDepot !== "All" && i.depot !== deliveryFilterDepot) return false;
    if (deliveryFilterStatus !== "All" && deriveStatus(i) !== deliveryFilterStatus) return false;
    if (!deliverySearch.trim()) return true;
    const q = deliverySearch.toLowerCase();
    return (
      i.client?.toLowerCase().includes(q) || i.project?.toLowerCase().includes(q) ||
      i.constructionSite?.toLowerCase().includes(q) || i.unitCode?.toLowerCase().includes(q) ||
      i.jobNumber?.toLowerCase().includes(q) || i.id?.toLowerCase().includes(q) ||
      i.shkNumber?.toLowerCase().includes(q) ||
      (i.packages || []).some((p) => (p.code || "").toLowerCase().includes(q))
    );
  });
  const pending = activeItemsList.filter((i) => deriveStatus(i) === "pending_collection");
  const billable = openForDelivery.filter((i) => storageInfo(i)?.billable);
  const lfdWarnings = activeItemsList.filter((i) => { const a = lfdAlert(i); return a && (a.level === "soon" || a.level === "overdue"); });
  // Reads a stack of packing lists in one go - Excel straight off the sheet, PDF through the
  // same scanner the single-file screen uses - and writes them out as the summary sheet the
  // packing list importer reads back in. Desktop work: a hundred files at a time is not
  // something anyone does on a phone.
  const [plrRows, setPlrRows] = useState([]);
  const [plrSaved, setPlrSaved] = useState("");
  // Reading a stack of PDFs takes minutes and the corrections take longer, so the table is
  // kept and restored rather than lost to a closed tab. Saved on demand, not on every
  // keystroke - a save that fires while someone is typing is a save nobody trusts.
  useEffect(() => {
    let live = true;
    storageGet("packingListReaderRows").then((v) => {
      if (!live || !v) return;
      try {
        const parsed = JSON.parse(v);
        if (Array.isArray(parsed.rows) && parsed.rows.length) {
          setPlrRows(withConflicts(parsed.rows));
          setPlrSaved(t.plrRestored(parsed.rows.length, parsed.at || ""));
        }
      } catch (err) { /* nothing usable stored */ }
    }).catch(() => {});
    return () => { live = false; };
  }, []);
  async function savePackingListRows() {
    try {
      await storageSet("packingListReaderRows", JSON.stringify({ rows: plrRows, at: todayStr() }));
      setPlrSaved(t.plrSavedNote(plrRows.length));
    } catch (err) {
      setPlrSaved(t.plrSaveFailed(String((err && err.message) || err)));
    }
  }
  const [plrBusy, setPlrBusy] = useState("");
  const [plrNotes, setPlrNotes] = useState([]);
  const plrInputRef = useRef(null);
  async function handlePackingListReaderFiles(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    if (!files.length) return;
    const rows = [];
    const notes = [];
    for (const file of files) {
      setPlrBusy(t.plrReading(file.name));
      try {
        if (/\.pdf$/i.test(file.name)) {
          if (file.size > 3.2 * 1024 * 1024) {
            notes.push(`${file.name}: ${t.plrTooBig(Math.round(file.size / 1024))}`);
            continue;
          }
          const base64 = await new Promise((res, rej) => {
            const fr = new FileReader();
            fr.onerror = () => rej(new Error("could not be read"));
            fr.onload = () => res(String(fr.result).split(",")[1]);
            fr.readAsDataURL(file);
          });
          const parsed = await postPdfScan({
            model: "claude-sonnet-4-6", max_tokens: 16000,
            messages: [{ role: "user", content: [
              { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
              { type: "text", text: PDF_SCAN_PROMPT },
            ] }],
          });
          const reconciled = reconcileGroupCaseCodes(parsed.groups || [], parsed.shippingMarks);
          if (reconciled.corrections.length) {
            notes.push(`${file.name}: ${reconciled.corrections.map((c) => `${c.lot} (${c.changed.join(", ")})`).join("; ")}`);
          }
          (reconciled.groups || []).forEach((g) => {
            const packages = (g.packages || []).length
              ? g.packages
              : (g.caseNumbers || []).flatMap((m) => String(m || "").split(",")).map((c) => ({ code: c.trim() })).filter((p) => p.code);
            rows.push(packingListSummaryRow(file.name, parsed.client, parsed.project, g.lot,
              packages,
              Number(g.statedWeightKg) || packages.reduce((n, p) => n + (Number(p.weightKg) || 0), 0),
              Number(g.statedCbm) || packages.reduce((n, p) => n + (Number(p.cbm) || 0), 0),
              Number(g.statedPackages) || 0, directory));
          });
        } else {
          const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, bookFiles: true });
          const { groups, client, project } = parsePackingListWorkbook(wb) || { groups: [] };
          if (!groups.length) { notes.push(`${file.name}: ${t.packingListNoStructure}`); continue; }
          groups.forEach((g) => rows.push(packingListSummaryRow(file.name, g.client || client, g.project || project,
            g.lot, g.packages, g.totalWeight, g.totalCbm, 0, directory)));
        }
      } catch (err) {
        notes.push(`${file.name}: ${String((err && err.message) || err)}`);
      }
    }
    setPlrBusy("");
    setPlrRows((prev) => withConflicts([...prev, ...rows]));
    setPlrNotes((prev) => [...prev, ...notes]);
  }
  // The same DM arriving on several files is normal - a consignment is routinely split
  // across two packing lists, and the machine bases or the car tops come on their own. What
  // is not normal is the same case appearing twice, or one DM being given two different
  // clients or sites. Those are marked so they can be sorted out in the spreadsheet before
  // any of it reaches the depot.
  function withConflicts(input) {
    // Repaired here as well as at scan time, so a row restored from a saved table, pasted
    // in, or typed by hand gets the same treatment as a fresh one. Without this, rows read
    // by an older build keep "B11 23" for ever, because nothing re-reads them.
    const rows = (input || []).map((r) => {
      const hint = /^\d{1,2}$/.test(String(r["DM or SHK or other Client Reference"] || "").trim())
        ? String(r["DM or SHK or other Client Reference"]).trim() : "";
      const fixed = String(r.Cases || "").split(",").map((c) => repairLiftFirstMarking(c, hint)).filter(Boolean).join(", ");
      return fixed === r.Cases ? r : { ...r, Cases: fixed };
    });
    const seenCase = new Map();
    const byRef = new Map();
    rows.forEach((r, i) => {
      const ref = String(r["DM or SHK or other Client Reference"] || "").trim().toUpperCase();
      // A case number repeated under a DIFFERENT reference is not a duplicate. Makers reuse
      // markings between consignments - 23D5423 belongs to both 13-DM-25-0625 and
      // 13-DM-26-0060 - and flagging those buried the real duplicates in noise. Only the
      // same case twice under the same reference is a problem.
      String(r.Cases || "").split(",").map((c) => c.trim()).filter(Boolean).forEach((c) => {
        const k = `${ref}\u0000${c.toUpperCase()}`;
        seenCase.set(k, [...(seenCase.get(k) || []), i]);
      });
      if (ref) byRef.set(ref, [...(byRef.get(ref) || []), i]);
    });
    const notes = rows.map(() => []);
    seenCase.forEach((idxs, key) => {
      if (idxs.length < 2) return;
      const code = key.split("\u0000")[1];
      const files = [...new Set(idxs.map((j) => rows[j]["File Name"]))];
      // Twice on the one file is worth saying differently from twice across two.
      idxs.forEach((i) => notes[i].push(t.plrDupCase(code, files.join(", "))));
    });
    byRef.forEach((idxs) => {
      if (idxs.length < 2) return;
      const clients = [...new Set(idxs.map((i) => rows[i].Client).filter(Boolean))];
      const sites = [...new Set(idxs.map((i) => rows[i].Project).filter(Boolean))];
      if (clients.length > 1) idxs.forEach((i) => notes[i].push(t.plrRefClientClash(clients.join(" / "))));
      if (sites.length > 1) idxs.forEach((i) => notes[i].push(t.plrRefSiteClash(sites.join(" / "))));
    });
    return rows.map((r, i) => {
      // Rebuilt, never appended to.
      const base = r.__baseCheck !== undefined ? r.__baseCheck : "";
      const all = [...new Set([base, ...notes[i]].filter(Boolean))];
      return { ...r, Check: all.join(" \u00b7 "), __conflict: notes[i].length > 0 };
    });
  }
  function exportPackingListSummary() {
    // Conflicting rows are put at the top rather than coloured. The spreadsheet writer in
    // this app cannot set a font colour - it writes no styles at all - so red would simply
    // not survive the download. Sorting them to the front, with the reason spelled out in
    // Check, puts them where they cannot be scrolled past instead.
    const ordered = [...plrRows].sort((a, b) => (b.__conflict ? 1 : 0) - (a.__conflict ? 1 : 0));
    const data = ordered.map((r) => PL_SUMMARY_COLUMNS.map((c) => (r[c] === undefined ? "" : r[c])));
    const ws = XLSX.utils.aoa_to_sheet([PL_SUMMARY_COLUMNS, ...data]);
    ws["!cols"] = PL_SUMMARY_COLUMNS.map((c) => ({ wch: c === "Cases" ? 70 : c === "File Name" || c === "Project" ? 28 : 16 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Packing Lists");
    XLSX.writeFile(wb, `packing-lists-${todayStr()}.xlsx`);
  }
  // Writes a date onto the record the ledger row came from. An arrival batch that has one
  // takes it; an entry with no batches at all has never had an arrival date, so the date
  // belongs on the entry itself, which is also what moves it out of "pending collection".
  function setMovementDate(move, date) {
    if (!date) return;
    persist(items.map((it) => {
      if (it.id !== move.entryId) return it;
      if (move.dir === "OUT") {
        return recomputeItemTotals({ ...it,
          deliveries: (it.deliveries || []).map((d) => (d.id === move.deliveryId ? { ...d, date } : d)) });
      }
      if (move.arrivalId) {
        const arrivals = (it.arrivals || []).map((a) => (a.id === move.arrivalId ? { ...a, date } : a));
        const dates = arrivals.map((a) => a.date).filter(Boolean).sort();
        return recomputeItemTotals({ ...it, arrivals, depotArrivalDate: dates[0] || date });
      }
      // No batch to hang it on: this is the entry's own arrival date.
      return recomputeItemTotals({ ...it, depotArrivalDate: date, awaitingCollection: false });
    }));
  }
  const ledgerAll = useMemo(() => storageLedger(activeItemsList, directory), [activeItemsList, directory]);
  const [ledgerOpen, setLedgerOpen] = useState({});
  const [ledgerAsOf, setLedgerAsOf] = useState(todayStr());
  // The ledger as it stood on a chosen day. Movements dated after it are set aside and the
  // running balance rebuilt without them, which answers "what was in the warehouse on the
  // 31st" for a billing month, and - with today's date, which is where it starts - keeps
  // paperwork entered ahead of time out of the figure. An undated movement has no place on
  // a dated view, so it is set aside too and counted separately rather than assumed recent.
  const ledgerSites = useMemo(() => {
    const cut = String(ledgerAsOf || "").trim();
    if (!cut) return ledgerAll.map((g) => ({ ...g, undatedMoves: [], later: 0, undated: 0 }));
    return ledgerAll.map((g) => {
      const kept = [];
      const undatedMoves = [];
      let later = 0;
      g.moves.forEach((m) => {
        // Kept, not discarded. A movement with no date cannot be placed on a dated view, but
        // hiding it is how it stays undated: it has to be visible, and obviously wrong, for
        // anyone to go and fix it.
        if (!m.date) { undatedMoves.push({ ...m }); return; }
        if (String(m.date) > cut) { later += 1; return; }
        kept.push({ ...m });
      });
      const undated = undatedMoves.length;
      let bal = 0;
      kept.forEach((m) => { bal += m.dir === "IN" ? m.pkgs : -m.pkgs; m.balance = bal; });
      return {
        ...g, moves: kept, undatedMoves, balance: bal, later, undated,
        inTotal: kept.filter((m) => m.dir === "IN").reduce((n, m) => n + m.pkgs, 0),
        outTotal: kept.filter((m) => m.dir === "OUT").reduce((n, m) => n + m.pkgs, 0),
      };
    }).filter((g) => g.moves.length || g.later || g.undated)
      .sort((a, b) => b.balance - a.balance);
  }, [ledgerAll, ledgerAsOf]);
  const ledgerTotal = useMemo(() => ledgerSites.reduce((n, g) => n + g.balance, 0), [ledgerSites]);
  const ledgerUndatedTotal = useMemo(() => ledgerSites.reduce((n, g) => n + (g.undated || 0), 0), [ledgerSites]);
  function exportLedger() {
    // Exports what is on screen, which is the ledger as at the chosen day - not everything
    // on file. A month-end figure that quietly included next month's paperwork would be
    // worse than useless.
    const head = ["Site", "Client", "FC Job No.", "In / Out", "Type", "Date", "Reference (DM / SHK)", "Lift", "Pkgs In", "Pkgs Out", "Balance"];
    const rows = [head];
    ledgerSites.forEach((g) => (g.undatedMoves || []).forEach((m) => rows.push([
      g.label, g.client, m.jobNumber, m.dir, m.type, "NO DATE", m.ref, m.lot,
      m.dir === "IN" ? m.pkgs : "", m.dir === "OUT" ? m.pkgs : "", "not counted",
    ])));
    ledgerSites.forEach((g) => g.moves.forEach((m) => rows.push([
      g.label, g.client, m.jobNumber, m.dir, m.type, m.date, m.ref, m.lot,
      m.dir === "IN" ? m.pkgs : "", m.dir === "OUT" ? m.pkgs : "", m.balance,
    ])));
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = head.map((h) => ({ wch: h === "Site" ? 34 : h === "Reference (DM / SHK)" ? 24 : 13 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Storage Ledger");
    XLSX.writeFile(wb, `farspeed-storage-ledger-as-at-${ledgerAsOf || todayStr()}.xlsx`);
  }
  const [storeMoves, setStoreMoves] = useState(null);
  const [storeName, setStoreName] = useState("");
  const storeInputRef = useRef(null);
  async function handleStoreListFile(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
      const moves = parseClientStoreList(wb);
      setStoreMoves(moves);
      setStoreName(file.name);
    } catch (err) {
      alert(t.storeListUnreadable(String((err && err.message) || err)));
    }
  }
  const storeRec = useMemo(
    () => (storeMoves ? reconcileStoreList(storeMoves, activeItemsList) : []),
    [storeMoves, activeItemsList]);
  const [checkInSearch, setCheckInSearch] = useState("");
  const [checkInOpen, setCheckInOpen] = useState({});
  const checkInGroups = useMemo(() => checkInAudit(activeItemsList), [activeItemsList]);
  // Searching looks at everything on a row that someone might have written down - the entry
  // id, the shipment it came in from, the lift or order code, the reference. Looking for
  // "INC-0559" should land on it whichever of those you happen to know.
  const checkInFiltered = useMemo(() => {
    const q = checkInSearch.trim().toLowerCase();
    if (!q) return checkInGroups;
    return checkInGroups
      .map((g) => ({ ...g, rows: g.rows.filter((r) => [g.ref, r.item.id, r.item.unitCode, r.source, r.date]
        .some((v) => String(v || "").toLowerCase().includes(q))) }))
      .filter((g) => g.rows.length);
  }, [checkInGroups, checkInSearch]);
  const overCheckedIn = useMemo(
    () => checkInGroups.reduce((n, g) => n + ((g.over || 0) > 0 ? g.over : 0), 0),
    [checkInGroups]);
  const duplicateGroups = useMemo(() => findDuplicateGroups(activeItemsList), [activeItemsList]);
  const duplicateIds = useMemo(() => new Set(duplicateGroups.flat().map((i) => i.id)), [duplicateGroups]);
  const siteTotals = useMemo(() => {
    const map = {};
    openForDelivery.forEach((it) => {
      const key = it.directoryId || sigPart(it.project) || "unspecified";
      if (!map[key]) {
        const dirEntry = it.directoryId ? (directory || []).find((d) => d.id === it.directoryId) : null;
        map[key] = {
          key,
          label: dirEntry ? dirEntry.siteEn : (it.project || "—"),
          labelZh: dirEntry ? dirEntry.siteZh : (it.constructionSite || ""),
          client: it.client,
          cbm: 0,
          kg: 0,
          pkgs: 0,
        };
      }
      map[key].cbm += remainingVolumeCbm(it);
      map[key].kg += remainingWeightKg(it);
      // Cases still at the depot, to match the cbm and kg beside it. Counting entries here
      // instead read "4" for a site holding sixteen cases across four entries.
      map[key].pkgs += Math.max(0, totalUnits(it) - deliveredUnits(it));
      // What is left at this site, broken down by the reference each lot came in under -
      // the DM number for Mitsubishi, the SHK for Schindler, whatever the maker uses. The
      // site row answers "how much"; this answers "which", which is the question that
      // actually gets asked when someone rings up about a job.
      const ref = String(it.invoiceNumber || "").trim() || String(it.unitCode || "").trim() || "\u2014";
      const byRef = (map[key].refs = map[key].refs || {});
      const b = (byRef[ref] = byRef[ref] || { ref, pkgs: 0, kg: 0, cbm: 0, codes: [], lots: [] });
      const left = remainingPackages(it);
      b.pkgs += left.length;
      b.kg += left.reduce((n, p) => n + (Number(p.weightKg) || 0), 0);
      b.cbm += left.reduce((n, p) => n + (Number(p.cbm) || 0), 0);
      b.codes.push(...left.map((p) => p.code));
      if (it.unitCode && !b.lots.includes(it.unitCode)) b.lots.push(it.unitCode);
    });
    // The last CFS, Devan and delivery each site saw, so a query about a site can be
    // answered from this table instead of opening every entry under it. Taken from every
    // entry at the site rather than only those still open, or a site whose last delivery
    // emptied an entry would show no delivery at all.
    const latest = {};
    (items || []).forEach((it) => {
      if (it.cancelled) return;
      const key = it.directoryId || sigPart(it.project) || "unspecified";
      if (!map[key]) return;
      if (!latest[key]) latest[key] = {};
      const note = (kind, date, ref) => {
        if (!date) return;
        const cur = latest[key][kind];
        if (!cur || String(date) > String(cur.date)) latest[key][kind] = { date, ref: ref || "" };
      };
      activeArrivals(it).forEach((a) => {
        // Three arrival types now, so a Return is not quietly filed under Devan.
        const at = String(a.type || it.arrivingType || "").toUpperCase();
        const kind = at === "CFS" ? "cfs" : at === "RETURN" ? "return" : "devan";
        note(kind, a.date, a.declaredSource || it.jobNumber);
      });
      // An entry carrying no arrival batches still records when and how it came in.
      if (!activeArrivals(it).length) {
        const at0 = String(it.arrivingType || "").toUpperCase();
        note(at0 === "CFS" ? "cfs" : at0 === "RETURN" ? "return" : "devan", it.depotArrivalDate, it.jobNumber);
      }
      activeDeliveries(it).forEach((d) => note("delivery", d.date, d.jobNumber));
    });
    return Object.values(map)
      .map((s) => ({
        ...s, cbm: Math.round(s.cbm * 1000) / 1000, kg: Math.round(s.kg * 10) / 10,
        latest: latest[s.key] || {},
        refs: Object.values(s.refs || {})
          .map((b) => ({ ...b, kg: Math.round(b.kg * 10) / 10, cbm: Math.round(b.cbm * 1000) / 1000 }))
          .filter((b) => b.pkgs > 0)
          .sort((a, b) => b.pkgs - a.pkgs),
      }))
      .sort((a, b) => b.cbm - a.cbm);
  }, [openForDelivery, items, directory]);
  const jobLog = useMemo(() => {
    const rows = [];
    items.forEach((it) => {
      if (it.jobNumber && !it.cancelled) {
        rows.push({
          jobNumber: it.jobNumber,
          type: it.arrivingType,
          date: it.depotArrivalDate,
          client: it.client,
          site: it.project,
          recordedBy: it.recordedBy,
          sheet: { type: it.arrivingType, item: it },
        });
      }
      (it.deliveries || []).forEach((d) => {
        if (d.jobNumber && !d.cancelled) {
          rows.push({
            jobNumber: d.jobNumber,
            type: "Delivery",
            date: d.date,
            client: it.client,
            site: it.project,
            recordedBy: d.recordedBy,
            sheet: { type: "Delivery", item: it, delivery: d },
          });
        }
      });
    });
    // Job numbers are issued in sequence, so sorting them numerically keeps 2608099 before
    // 2608100 rather than after it, the way a plain string sort would have it.
    return rows.sort((a, b) => b.jobNumber.localeCompare(a.jobNumber, undefined, { numeric: true })
      || String(b.date || "").localeCompare(String(a.date || "")));
  }, [items]);

  // Job Log filters. Options come from the log itself, so a filter can never come back
  // empty, and job numbers sort in sequence - the order they were issued in, which is how
  // a gap or a reuse shows itself.
  const [jobLogSearch, setJobLogSearch] = useState("");
  const [jobLogType, setJobLogType] = useState("All");
  const [jobLogClient, setJobLogClient] = useState("All");
  const [jobLogSite, setJobLogSite] = useState("All");
  const [jobLogRecordedBy, setJobLogRecordedBy] = useState("All");
  const [jobLogFrom, setJobLogFrom] = useState("");
  const [jobLogTo, setJobLogTo] = useState("");
  const jobLogOptions = useMemo(() => {
    const grab = (key) => [...new Set(jobLog.map((r) => String(r[key] || "").trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    return { clients: grab("client"), sites: grab("site"), recordedBy: grab("recordedBy") };
  }, [jobLog]);
  const jobLogFiltered = useMemo(() => jobLog.filter((r) => {
    if (jobLogType !== "All" && r.type !== jobLogType) return false;
    if (jobLogClient !== "All" && String(r.client || "").trim() !== jobLogClient) return false;
    if (jobLogSite !== "All" && String(r.site || "").trim() !== jobLogSite) return false;
    if (jobLogRecordedBy !== "All" && String(r.recordedBy || "").trim() !== jobLogRecordedBy) return false;
    if (jobLogFrom && (!r.date || r.date < jobLogFrom)) return false;
    if (jobLogTo && (!r.date || r.date > jobLogTo)) return false;
    if (!jobLogSearch.trim()) return true;
    const q = jobLogSearch.toLowerCase();
    return [r.jobNumber, r.type, r.client, r.site, r.recordedBy,
      r.sheet && r.sheet.item ? r.sheet.item.jobRef : "",
      r.sheet && r.sheet.item ? r.sheet.item.unitCode : "",
      r.sheet && r.sheet.item ? r.sheet.item.shkNumber : "",
      r.sheet && r.sheet.item ? r.sheet.item.constructionSite : "",
    ].some((v) => String(v || "").toLowerCase().includes(q));
  }), [jobLog, jobLogSearch, jobLogType, jobLogClient, jobLogSite, jobLogRecordedBy, jobLogFrom, jobLogTo]);
  const jobLogFiltersOn = [jobLogType, jobLogClient, jobLogSite, jobLogRecordedBy].some((v) => v !== "All")
    || !!jobLogSearch.trim() || !!jobLogFrom || !!jobLogTo;
  function clearJobLogFilters() {
    setJobLogSearch(""); setJobLogType("All"); setJobLogClient("All");
    setJobLogSite("All"); setJobLogRecordedBy("All"); setJobLogFrom(""); setJobLogTo("");
  }

  const cancelledJobs = useMemo(() => {
    const rows = [];
    items.forEach((it) => {
      if (it.jobNumber && it.cancelled) {
        rows.push({
          jobNumber: it.jobNumber,
          type: it.arrivingType,
          date: it.depotArrivalDate,
          client: it.client,
          site: it.project,
          recordedBy: it.recordedBy,
          sheet: { type: it.arrivingType, item: it },
          onRestore: () => handleRestoreItem(it.id),
          onPurge: () => handlePermanentlyDeleteItem(it.id),
        });
      }
      (it.deliveries || []).forEach((d) => {
        if (d.jobNumber && d.cancelled) {
          rows.push({
            jobNumber: d.jobNumber,
            type: "Delivery",
            date: d.date,
            client: it.client,
            site: it.project,
            recordedBy: d.recordedBy,
            sheet: { type: "Delivery", item: it, delivery: d },
            onRestore: () => handleRestoreDelivery(it.id, d.id),
            onPurge: () => handlePermanentlyDeleteDelivery(it.id, d.id),
          });
        }
      });
    });
    return rows.sort((a, b) => b.jobNumber.localeCompare(a.jobNumber));
  }, [items]);

  if (authUser === undefined) {
    return <div className="p-8 text-sm" style={{ color: colors.inkFaint }}>{t.loadingMsg}</div>;
  }
  if (authUser === null) {
    return <LoginScreen onLoggedIn={handleLoggedIn} colors={colors} t={t} lang={lang} />;
  }

  if (!loaded) {
    return <div className="p-8 text-sm" style={{ color: colors.inkFaint }}>{t.loadingMsg}</div>;
  }

  return (
    <div style={{ background: colors.bg, minHeight: "100%", fontFamily: FONT_BODY }} className="w-full">
      <style>{FONT_IMPORT}</style>

      <div style={{ background: colors.navy }} className="px-4 md:px-6 py-3 md:py-4 flex items-center justify-between gap-3">
        <div className="min-w-0">
          <div className="text-lg md:text-2xl font-bold tracking-tight truncate" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>FARSPEED CONTRACTORS LTD</div>
          <div className="text-[10px] md:text-xs tracking-widest uppercase" style={{ color: colors.amber }}>{t.appSubtitle}</div>
        </div>
        <button
          className="md:hidden flex items-center justify-center rounded"
          style={{ width: 40, height: 40, background: colors.navySoft, color: colors.onDark, fontSize: 20, flexShrink: 0 }}
          onClick={() => setMobileMenuOpen((o) => !o)}
          aria-label={t.mobileMenuLabel}
        >
          {mobileMenuOpen ? "\u2715" : "\u2630"}
        </button>
        <div className="hidden md:flex items-center gap-3 flex-wrap">
          <div className="flex gap-1 rounded-lg p-1" style={{ background: colors.navySoft }}>
            {[
              ["dashboard", t.navDashboard],
              ["inventory", t.navInventory],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setNewEntryMenuOpen(false); setSettingsOpen(false); setView(k); }}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark }}>
                {label}
              </button>
            ))}

            <button
              onClick={() => { setEditing(null); setSettingsOpen(false); setView("add"); }}
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, background: view === "add" ? colors.amber : "transparent", color: view === "add" ? colors.ink : colors.onDark }}
            >
              {t.navNewEntry}
            </button>

            {[
              ["upload", t.navUpload],
              ["incoming", t.navIncoming],
              ["billing", t.navBilling],
              ["directory", t.navDirectory],
              ["joblog", t.navJobLog],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setNewEntryMenuOpen(false); setSettingsOpen(false); setView(k); }}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark }}>
                {label}
              </button>
            ))}

            <div className="relative">
              <button
                onClick={() => { setSettingsOpen((o) => !o); setNewEntryMenuOpen(false); }}
                title={t.settingsLabel}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: ["duplicates", "cancelledjobs", "checkins", "ledger", "plreader"].includes(view) ? colors.amber : "transparent", color: ["duplicates", "cancelledjobs", "checkins", "ledger", "plreader"].includes(view) ? colors.ink : colors.onDark }}
              >
                ⚙
              </button>
              {settingsOpen && (
                <div className="absolute right-0 mt-1 rounded-lg overflow-hidden z-20" style={{ background: colors.surface, border: `1px solid ${colors.line}`, minWidth: 200 }}>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}
                    onClick={() => { setView("plreader"); setSettingsOpen(false); }}
                  >
                    {t.navPlReader}
                  </button>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY, borderTop: `1px solid ${colors.surfaceDim}` }}
                    onClick={() => { setView("ledger"); setSettingsOpen(false); }}
                  >
                    {t.navLedger}
                  </button>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: overCheckedIn > 0 ? colors.red : colors.ink, fontFamily: FONT_DISPLAY, borderTop: `1px solid ${colors.surfaceDim}` }}
                    onClick={() => { setView("checkins"); setSettingsOpen(false); }}
                  >
                    {overCheckedIn > 0 ? t.navCheckInsCount(overCheckedIn) : t.navCheckIns}
                  </button>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY, borderTop: `1px solid ${colors.surfaceDim}` }}
                    onClick={() => { setView("duplicates"); setSettingsOpen(false); }}
                  >
                    {duplicateGroups.length > 0 ? t.navDuplicatesCount(duplicateGroups.length) : t.navDuplicatesShort}
                  </button>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY, borderTop: `1px solid ${colors.surfaceDim}` }}
                    onClick={() => { setView("cancelledjobs"); setSettingsOpen(false); }}
                  >
                    {cancelledJobs.length > 0 ? `${t.navCancelledJobs} (${cancelledJobs.length})` : t.navCancelledJobs}
                  </button>
                </div>
              )}
            </div>
          </div>
          <div className="flex items-center gap-2 rounded-lg px-2 py-1" style={{ background: colors.navySoft }}>
            <span className="text-xs" style={{ color: colors.onDark, opacity: 0.7, fontFamily: FONT_DISPLAY }}>{t.signedInAs}:</span>
            <span className="text-sm font-semibold" style={{ color: colors.amber, fontFamily: FONT_DISPLAY }}>{authUser}</span>
            <button className="text-xs font-semibold underline" style={{ color: colors.onDark, opacity: 0.8 }} onClick={() => setChangePasswordOpen(true)}>{t.changePasswordLink}</button>
            <button className="text-xs font-semibold underline" style={{ color: colors.onDark, opacity: 0.8 }} onClick={handleLogout}>{t.logoutBtn}</button>
          </div>
          <div className="flex gap-1 rounded-lg p-1" style={{ background: colors.navySoft }}>
            <button
              title={t.langToggleLabel}
              onClick={() => setLang((l) => (l === "en" ? "zh" : "en"))}
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark }}
            >
              {lang === "en" ? "中文" : "EN"}
            </button>
            <button
              title={t.themeToggleLabel}
              onClick={() => setTheme((th) => (th === "light" ? "dark" : "light"))}
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark }}
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </div>
      </div>

      {mobileMenuOpen && (
        <div className="md:hidden flex flex-col" style={{ background: colors.navySoft, borderTop: `1px solid ${colors.navy}` }}>
          {[
            ["dashboard", t.navDashboard],
            ["inventory", t.navInventory],
            ["add", t.navNewEntry],
            ["upload", t.navUpload],
            ["incoming", t.navIncoming],
            ["billing", t.navBilling],
            ["directory", t.navDirectory],
            ["joblog", t.navJobLog],
            ["plreader", t.navPlReader],
            ["ledger", t.navLedger],
            ["checkins", overCheckedIn > 0 ? t.navCheckInsCount(overCheckedIn) : t.navCheckIns],
            ["duplicates", duplicateGroups.length > 0 ? t.navDuplicatesCount(duplicateGroups.length) : t.navDuplicatesShort],
            ["cancelledjobs", cancelledJobs.length > 0 ? `${t.navCancelledJobs} (${cancelledJobs.length})` : t.navCancelledJobs],
          ].map(([k, label]) => (
            <button
              key={k}
              onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setNewEntryMenuOpen(false); setSettingsOpen(false); setMobileMenuOpen(false); setView(k); }}
              className="text-left px-5 py-3 text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark, borderBottom: `1px solid ${colors.navy}` }}
            >
              {label}
            </button>
          ))}
          <div className="flex items-center justify-between px-5 py-3" style={{ borderBottom: `1px solid ${colors.navy}` }}>
            <span className="text-xs" style={{ color: colors.onDark, opacity: 0.7, fontFamily: FONT_DISPLAY }}>{t.signedInAs}: <span className="font-semibold" style={{ color: colors.amber }}>{authUser}</span></span>
          </div>
          <button
            className="text-left px-5 py-3 text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, borderBottom: `1px solid ${colors.navy}` }}
            onClick={() => { setChangePasswordOpen(true); setMobileMenuOpen(false); }}
          >
            {t.changePasswordLink}
          </button>
          <button
            className="text-left px-5 py-3 text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, borderBottom: `1px solid ${colors.navy}` }}
            onClick={handleLogout}
          >
            {t.logoutBtn}
          </button>
          <div className="flex gap-2 px-5 py-3">
            <button
              className="flex-1 px-3 py-2 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, background: colors.navy }}
              onClick={() => setLang((l) => (l === "en" ? "zh" : "en"))}
            >
              {lang === "en" ? "中文" : "EN"}
            </button>
            <button
              className="flex-1 px-3 py-2 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, background: colors.navy }}
              onClick={() => setTheme((th) => (th === "light" ? "dark" : "light"))}
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </div>
      )}

      {/* The reader and the ledger are wide tables - eleven columns of case markings - and a
          six-column page frame squeezed them into a ribbon down the middle of a monitor.
          Those two get the full width; everything else keeps the narrower measure, which is
          easier to read for forms and lists. */}
      <div className={`p-3 md:p-6 mx-auto${["plreader", "ledger", "checkins"].includes(view) ? " w-full" : " max-w-6xl"}`}>
        {error && <div className="mb-4 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>}
        {conflictKey && (
          <div className="mb-4 px-3 py-3 rounded text-sm" style={{ background: colors.amberSoft, color: colors.ink, borderLeft: `3px solid ${colors.amber}` }}>
            <div className="font-semibold" style={{ fontFamily: FONT_DISPLAY }}>{t.conflictTitle}</div>
            <div className="mt-1">{t.conflictBody}</div>
            <button className="mt-2 text-xs font-semibold underline" onClick={() => setConflictKey("")}>{t.conflictDismiss}</button>
          </div>
        )}

        {view === "dashboard" && (
          <div className="flex flex-col gap-6">
            <div className="flex justify-end">
              <button
                className="text-xs font-semibold underline"
                style={{ color: colors.inkFaint }}
                onClick={handleResetDeliveries}
              >
                {t.resetBtn}
              </button>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <StatCard label={t.statAtDepot} value={openForDelivery.length} colors={colors} />
              <StatCard label={t.statPending} value={pending.length} tone="grey" colors={colors} />
              <StatCard label={t.statBillable} value={billable.length} tone="red" colors={colors} />
              <StatCard label={t.statLfd} value={lfdWarnings.length} tone="amber" colors={colors} />
            </div>

            {duplicateGroups.length > 0 && (
              <div className="rounded-lg p-4 flex items-center justify-between flex-wrap gap-2" style={{ background: colors.amberSoft }}>
                <span className="text-sm" style={{ color: colors.amberText }}>{t.dupBanner(duplicateGroups.length)}</span>
                <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => setView("duplicates")}>
                  {t.reviewDuplicatesBtn}
                </button>
              </div>
            )}

            {lfdWarnings.length > 0 && (
              <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
                <h3 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.red }}>{t.lfdSectionTitle}</h3>
                <div className="flex flex-col gap-2">
                  {lfdWarnings.map((i) => (
                    <div key={i.id} className="flex items-center justify-between text-sm py-1.5" style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                      <span style={{ fontFamily: FONT_MONO }}>{i.id}</span>
                      <span className="flex-1 mx-3 truncate">{i.client} · {i.project}</span>
                      <span className="mr-3" style={{ color: colors.inkFaint }}>LFD {fmt(i.terminalLFD)}</span>
                      <StatusBadge item={i} colors={colors} t={t} />
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.depotOverviewTitle}</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {DEPOTS.map((d) => {
                  const totals = depotRemainingTotals(activeItemsList, d);
                  return (
                    <div key={d} className="rounded-lg p-3" style={{ background: colors.surfaceDim }}>
                      <div className="text-sm font-semibold mb-2" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
                        {depotLabel(d, lang)} <span style={{ color: colors.inkFaint, fontWeight: 400 }}>({totals.count} {t.depotOverviewItemsLabel})</span>
                      </div>
                      <div className="flex gap-6">
                        <div>
                          <div className="text-2xl font-bold" style={{ color: colors.ink, fontFamily: FONT_MONO }}>{totals.cbm}</div>
                          <div className="text-xs" style={{ color: colors.inkFaint }}>{t.jsCbm}</div>
                        </div>
                        <div>
                          <div className="text-2xl font-bold" style={{ color: colors.ink, fontFamily: FONT_MONO }}>{totals.kg}</div>
                          <div className="text-xs" style={{ color: colors.inkFaint }}>{t.jsKgs}</div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.billableSectionTitle}</h3>
              {billable.length === 0 ? (
                <p className="text-sm" style={{ color: colors.inkFaint }}>{t.billableEmptyMsg}</p>
              ) : (
                <div className="flex flex-col gap-2">
                  {billable.map((i) => {
                    const info = storageInfo(i);
                    return (
                      <div key={i.id} className="flex items-center justify-between text-sm py-1.5" style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <span style={{ fontFamily: FONT_MONO }}>{i.id}</span>
                        <span className="flex-1 mx-3 truncate">{i.client} · {i.project}</span>
                        <span className="mr-3" style={{ color: colors.inkFaint }}>{t.sinceLabel} {fmt(info.freeUntil)}</span>
                        <StatusBadge item={i} colors={colors} t={t} />
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        )}

        {view === "inventory" && (
          <div className="flex flex-col gap-4">
            <div className="flex justify-end">
              <button
                onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setView("exit"); }}
                className="px-4 py-2 rounded text-sm font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              >
                {t.navDeliveries} &rarr;
              </button>
            </div>
            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <div
                className="px-4 py-2 flex items-center justify-between cursor-pointer"
                style={{ background: colors.surfaceDim }}
                onClick={() => setSiteTotalsOpen((o) => !o)}
              >
                <span className="text-sm font-bold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.siteTotalsTitle}</span>
                <span className="text-xs font-semibold" style={{ color: colors.amberText }}>
                  {siteTotalsOpen ? t.siteTotalsToggleHide : t.siteTotalsToggleShow}
                </span>
              </div>
              {siteTotalsOpen && (
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {[t.siteTotalsColSite, t.siteTotalsColClient, t.siteTotalsColPkgs, t.siteTotalsColCbm, t.siteTotalsColKg,
                        t.siteTotalsColLastCfs, t.siteTotalsColLastDevan, t.siteTotalsColLastReturn, t.siteTotalsColLastDelivery].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {siteTotals.length === 0 && (
                      <tr><td colSpan={9} className="px-3 py-4 text-center text-sm" style={{ color: colors.inkFaint }}>{t.siteTotalsNoneMsg}</td></tr>
                    )}
                    {siteTotals.map((s) => (
                      <React.Fragment key={s.key}>
                      <tr
                        style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: s.refs.length ? "pointer" : "default" }}
                        onClick={() => s.refs.length && setOpenSite((o) => ({ ...o, [s.key]: !o[s.key] }))}
                      >
                        <td className="px-3 py-2">
                          <div>{s.label}</div>
                          {s.labelZh && <div className="text-xs" style={{ color: colors.inkFaint }}>{s.labelZh}</div>}
                        </td>
                        <td className="px-3 py-2">{s.client}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.pkgs}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.cbm}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.kg}</td>
                        {["cfs", "devan", "return", "delivery"].map((kind) => (
                          <td key={kind} className="px-3 py-2 whitespace-nowrap">
                            {s.latest[kind] ? (
                              <>
                                <div>{fmt(s.latest[kind].date)}</div>
                                {s.latest[kind].ref && (
                                  <div className="text-xs truncate" style={{ color: colors.inkFaint, fontFamily: FONT_MONO, maxWidth: 170 }} title={s.latest[kind].ref}>
                                    {s.latest[kind].ref}
                                  </div>
                                )}
                              </>
                            ) : <span style={{ color: colors.inkFaint }}>—</span>}
                          </td>
                        ))}
                      </tr>
                      {openSite[s.key] && s.refs.length > 0 && (
                        <tr style={{ background: colors.surfaceDim }}>
                          <td colSpan={9} className="px-3 py-2">
                            <table className="w-full text-xs">
                              <thead>
                                <tr>
                                  {[t.siteRefsColRef, t.siteRefsColLots, t.siteTotalsColPkgs, t.siteTotalsColKg, t.siteTotalsColCbm, t.siteRefsColCases].map((h, hi) => (
                                    <th key={hi} className="text-left px-2 py-1 font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {s.refs.map((b) => (
                                  <tr key={b.ref} style={{ borderTop: `1px solid ${colors.line}` }}>
                                    <td className="px-2 py-1 font-semibold" style={{ color: colors.ink }}>{b.ref}</td>
                                    <td className="px-2 py-1" style={{ color: colors.inkFaint }}>{b.lots.join(", ") || "\u2014"}</td>
                                    <td className="px-2 py-1" style={{ fontFamily: FONT_MONO }}>{b.pkgs}</td>
                                    <td className="px-2 py-1" style={{ fontFamily: FONT_MONO }}>{b.kg}</td>
                                    <td className="px-2 py-1" style={{ fontFamily: FONT_MONO }}>{b.cbm}</td>
                                    <td className="px-2 py-1" style={{ color: colors.inkFaint, wordBreak: "break-word", maxWidth: 420 }} title={b.codes.join(", ")}>
                                      {b.codes.join(", ")}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            <div className="flex flex-wrap gap-3 items-end">
              <Field label={t.searchLabel} colors={colors}>
                <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }} placeholder={t.searchPlaceholder} value={search} onChange={(e) => setSearch(e.target.value)} />
              </Field>
              <Field label={t.clientLabel} colors={colors}>
                <select className={inputClass} style={inputStyleFor(colors)} value={filterClient} onChange={(e) => setFilterClient(e.target.value)}>
                  <option>All</option>
                  {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              <Field label={t.statusLabel} colors={colors}>
                <select className={inputClass} style={inputStyleFor(colors)} value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                  <option value="All">{t.statusAll}</option>
                  <option value="pending_collection">{t.statusPending}</option>
                  <option value="at_depot">{t.statusAtDepot}</option>
                  <option value="partial">{t.statusPartial}</option>
                  <option value="delivered">{t.statusDelivered}</option>
                </select>
              </Field>
              <Field label={t.depotLabel} colors={colors}>
                <select className={inputClass} style={inputStyleFor(colors)} value={filterDepot} onChange={(e) => setFilterDepot(e.target.value)}>
                  <option>All</option>
                  {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                </select>
              </Field>
              {(() => {
                const chosen = filtered.filter((i) => pickedRows[i.id]);
                if (!chosen.length) return null;
                const pkgs = chosen.reduce((n, i) => n + totalUnits(i), 0);
                const delivered = chosen.filter((i) => activeDeliveries(i).length).length;
                return (
                  <span className="flex items-center gap-3 ml-auto">
                    <span className="text-sm font-semibold" style={{ color: colors.ink }}>{t.invSelectedCount(chosen.length, pkgs)}</span>
                    <button className="text-sm font-semibold" style={{ color: colors.red }}
                      onClick={() => {
                        if (!window.confirm(t.invBulkDeleteConfirm(chosen.length, pkgs, delivered))) return;
                        handlePermanentlyDeleteItems(chosen.map((i) => i.id));
                        setPickedRows({});
                      }}>{t.invBulkDeleteBtn}</button>
                    <button className="text-sm" style={{ color: colors.inkFaint }} onClick={() => setPickedRows({})}>{t.legacyClearSelection}</button>
                  </span>
                );
              })()}
              <button className={`px-3 py-1.5 rounded text-sm font-semibold${Object.values(pickedRows).some(Boolean) ? "" : " ml-auto"}`} style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => exportToExcel(filtered)}>
                {t.exportBtn(filtered.length)}
              </button>
              <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => { setEditing(null); setView("add"); }}>
                {t.newEntryBtn}
              </button>
            </div>

            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <div className="overflow-x-auto">
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      <th className="px-3 py-2" style={{ width: 34 }}>
                        <input type="checkbox"
                          checked={filtered.length > 0 && filtered.every((i) => pickedRows[i.id])}
                          onChange={(e) => {
                            const next = { ...pickedRows };
                            filtered.forEach((i) => { if (e.target.checked) next[i.id] = true; else delete next[i.id]; });
                            setPickedRows(next);
                          }} />
                      </th>
                      {[t.colId, t.colJobNo, t.colClient, t.colProjectSite, t.colUnit, t.colDepot, t.colDepotArrival, t.colStatus, t.colPackages, t.colCbm, t.colKg, ""].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.length === 0 && (
                      <tr><td colSpan={13} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.noRecordsMsg}</td></tr>
                    )}
                    {filtered.map((i) => (
                      <React.Fragment key={i.id}>
                      <tr
                        style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer", background: pickedRows[i.id] ? colors.amberSoft : undefined }}
                        onClick={() => setExpandedRowId((prev) => (prev === i.id ? null : i.id))}
                      >
                        <td className="px-3 py-2">
                          <input type="checkbox" checked={!!pickedRows[i.id]}
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) => setPickedRows((p) => ({ ...p, [i.id]: e.target.checked }))} />
                        </td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{i.id}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{i.jobNumber || "—"}</td>
                        <td className="px-3 py-2">{i.client}</td>
                        <td className="px-3 py-2 max-w-[220px]">
                          <div className="truncate">{i.project}</div>
                          {i.constructionSite && <div className="truncate text-xs" style={{ color: colors.inkFaint }}>{i.constructionSite}</div>}
                        </td>
                        <td className="px-3 py-2">{i.unitCode || "—"}</td>
                        <td className="px-3 py-2 whitespace-nowrap">{depotDisplay(i.depot, lang)}</td>
                        <td className="px-3 py-2">{fmt(i.depotArrivalDate)}</td>
                        <td className="px-3 py-2">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <StatusBadge item={i} colors={colors} t={t} />
                            {duplicateIds.has(i.id) && <Badge tone="amber" colors={colors}>{t.duplicateBadge}</Badge>}
                          </div>
                        </td>
                        <td className="px-3 py-2">{totalUnits(i)}</td>
                        <td className="px-3 py-2">{i.volumeCbm || "—"}</td>
                        <td className="px-3 py-2">{i.weightKg || "—"}</td>
                        <td className="px-3 py-2" onClick={(e) => e.stopPropagation()}>
                          <div className="flex flex-wrap justify-end items-center gap-x-3 gap-y-1">
                            {["at_depot", "partial"].includes(deriveStatus(i)) && (
                              <button className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => { setExitingItems([i]); setView("exit"); }}>{t.deliverBtn}</button>
                            )}
                            <button className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => { setEditing(i); setView("add"); }}>{t.editBtn}</button>
                            {i.jobNumber && (
                              <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => handleCancelItem(i.id)}>{t.cancelJobBtn}</button>
                            )}
                            <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => handleDelete(i.id)}>{t.deleteBtn}</button>
                          </div>
                        </td>
                      </tr>
                      {expandedRowId === i.id && (
                        <tr style={{ background: colors.surfaceDim }}>
                          <td colSpan={12} className="px-4 py-3">
                            {(() => {
                              const remaining = remainingPackages(i);
                              const remKg = Math.round(remainingWeightKg(i) * 10) / 10;
                              const remCbm = Math.round(remainingVolumeCbm(i) * 1000) / 1000;
                              return (
                                <div className="flex flex-col gap-2">
                                  <div className="flex flex-wrap gap-4 text-sm" style={{ color: colors.ink }}>
                                    <span><strong>{remainingUnits(i)}</strong> / {totalUnits(i)} {t.jsPkgs} {t.inventoryRemainingLabel}</span>
                                    <span><strong>{remCbm || 0}</strong> {t.jsCbm} {t.inventoryRemainingLabel}</span>
                                    <span><strong>{remKg || 0}</strong> {t.jsKgs} {t.inventoryRemainingLabel}</span>
                                  </div>
                                  {i.packages && i.packages.length > 0 && (
                                    remaining.length > 0 ? (
                                      <div className="flex flex-wrap gap-1.5">
                                        {(() => {
                                          const awaiting = new Set(notYetArrivedPackages(i).map((p) => p.code));
                                          return remaining.map((p) => (
                                            awaiting.has(p.code) ? (
                                              <span key={p.code} className="px-2 py-1 rounded text-xs" style={{ background: colors.surfaceDim, border: `1px dashed ${colors.line}`, color: colors.inkFaint }} title={t.notYetArrivedHint}>
                                                {p.code} · {t.notYetArrivedTag}
                                              </span>
                                            ) : (
                                              <span key={p.code} className="px-2 py-1 rounded text-xs" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.ink }}>
                                                {p.code}{p.description ? ` — ${p.description}` : ""}
                                              </span>
                                            )
                                          ));
                                        })()}
                                      </div>
                                    ) : (
                                      <span className="text-xs" style={{ color: colors.inkFaint }}>{t.inventoryNoRemainingPkgsMsg}</span>
                                    )
                                  )}
                                  {activeDeliveries(i).length > 0 && (
                                    <div className="mt-1">
                                      <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{t.deliveryHistoryLabel}</div>
                                      <div className="flex flex-col gap-1.5">
                                        {[...activeDeliveries(i)].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((d) => (
                                          <div key={d.id} className="flex flex-wrap items-center gap-2 text-xs" style={{ color: colors.ink }}>
                                            <span className="font-semibold" style={{ fontFamily: FONT_MONO }}>{d.jobNumber || "—"}</span>
                                            <span>{fmt(d.date)}</span>
                                            <span style={{ color: colors.inkFaint }}>
                                              {d.codes && d.codes.length > 0 ? d.codes.join(", ") : `${d.packageCount || "?"} ${t.jsPkgs}`}
                                            </span>
                                            <SignedDocControl docKey={`signedDoc:${i.id}:${d.id}`} colors={colors} t={t} />
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  )}
                                  <div className="flex items-center gap-2 text-xs mt-1" style={{ color: colors.inkFaint }}>
                                    <span>{t.signedDocArrivalLabel}:</span>
                                    <SignedDocControl docKey={`signedDoc:${i.id}:arrival`} colors={colors} t={t} />
                                  </div>
                                </div>
                              );
                            })()}
                          </td>
                        </tr>
                      )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {view === "exit" && (
          <div className="flex flex-col gap-4">
            {exitingItems.length > 0 ? (
              <DeliveryForm deliveryItems={exitingItems} onAddDelivery={handleAddDelivery} onAddCombinedDelivery={handleAddCombinedDelivery} onDeleteDelivery={handleDeleteDelivery} onUpdateDelivery={handleUpdateDelivery}
                onCancel={() => { setExitingItems([]); setDeliveryPickerSelection([]); setView("inventory"); }} onPrintJobSheet={setPrintJobSheet}
                employees={employees} currentUser={currentUser} items={items} colors={colors} t={t} lang={lang} />
            ) : (
              <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-3 flex flex-wrap gap-3 items-end" style={{ background: colors.surfaceDim }}>
                  <Field label={t.searchLabel} colors={colors}>
                    <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }} placeholder={t.deliverySearchPlaceholder} value={deliverySearch} onChange={(e) => setDeliverySearch(e.target.value)} />
                  </Field>
                  <Field label={t.clientLabel} colors={colors}>
                    <select className={inputClass} style={inputStyleFor(colors)} value={deliveryFilterClient} onChange={(e) => setDeliveryFilterClient(e.target.value)}>
                      <option>All</option>
                      {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label={t.depotLabel} colors={colors}>
                    <select className={inputClass} style={inputStyleFor(colors)} value={deliveryFilterDepot} onChange={(e) => setDeliveryFilterDepot(e.target.value)}>
                      <option>All</option>
                      {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                    </select>
                  </Field>
                  <Field label={t.statusLabel} colors={colors}>
                    <select className={inputClass} style={inputStyleFor(colors)} value={deliveryFilterStatus} onChange={(e) => setDeliveryFilterStatus(e.target.value)}>
                      <option value="All">{t.statusAll}</option>
                      <option value="at_depot">{t.statusAtDepot}</option>
                      <option value="partial">{t.statusPartial}</option>
                    </select>
                  </Field>
                  <div className="ml-auto">
                    <button
                      className="text-sm font-semibold px-3 py-2 rounded"
                      style={{ background: deliveryPickerSelection.length > 0 ? colors.amber : colors.surfaceDim, color: deliveryPickerSelection.length > 0 ? colors.ink : colors.inkFaint, fontFamily: FONT_DISPLAY, cursor: deliveryPickerSelection.length > 0 ? "pointer" : "default" }}
                      disabled={deliveryPickerSelection.length === 0}
                      onClick={() => setExitingItems(filteredForDelivery.filter((i) => deliveryPickerSelection.includes(i.id)))}
                    >
                      {t.recordCombinedBtn(deliveryPickerSelection.length)}
                    </button>
                  </div>
                </div>
                {deliveryPickerSelection.length > 0 && (
                  <div className="px-4 py-2 text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
                    {t.combinedSelectionHint}
                  </div>
                )}
                <div className="overflow-x-auto">
                  <table className="w-full text-sm" style={{ background: colors.surface }}>
                    <thead>
                      <tr style={{ background: colors.surfaceDim }}>
                        <th className="px-3 py-2" style={{ width: 32 }}></th>
                        {[t.colClient, t.colProjectSite, t.colUnit, t.colDepot, t.colDepotArrival, t.colStatus, ""].map((h) => (
                          <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredForDelivery.length === 0 && (
                        <tr><td colSpan={8} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.nothingAtDepotMsg}</td></tr>
                      )}
                      {filteredForDelivery.map((i) => {
                        const checked = deliveryPickerSelection.includes(i.id);
                        return (
                        <tr key={i.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, background: checked ? colors.amberSoft : "transparent" }}>
                          <td className="px-3 py-2">
                            <input type="checkbox" checked={checked} onChange={() => setDeliveryPickerSelection((prev) => checked ? prev.filter((x) => x !== i.id) : [...prev, i.id])} />
                          </td>
                          <td className="px-3 py-2">{i.client}</td>
                          <td className="px-3 py-2 max-w-[220px] truncate">{i.project}</td>
                          <td className="px-3 py-2">{i.unitCode || "—"}</td>
                          <td className="px-3 py-2 whitespace-nowrap">{depotDisplay(i.depot, lang)}</td>
                          <td className="px-3 py-2">{fmt(i.depotArrivalDate)}</td>
                          <td className="px-3 py-2"><StatusBadge item={i} colors={colors} t={t} /></td>
                          <td className="px-3 py-2 text-right">
                            <button className="text-xs font-semibold px-2 py-1 rounded" style={{ background: colors.amber, color: colors.ink }} onClick={() => setExitingItems([i])}>
                              {t.recordDeliveryBtn}
                            </button>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {view === "plreader" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.plrTitle}</h3>
              <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.plrDesc}</p>
              <input ref={plrInputRef} type="file" multiple accept=".pdf,.xlsx,.xls,.xlsm" className="hidden" onChange={handlePackingListReaderFiles} />
              <button className="px-3 py-1.5 rounded text-sm font-semibold mr-2"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: plrBusy ? 0.6 : 1 }}
                disabled={!!plrBusy}
                onClick={() => plrInputRef.current && plrInputRef.current.click()}>
                {plrBusy || t.plrChooseBtn}
              </button>
              {plrRows.length > 0 && (
                <>
                  <button className="px-3 py-1.5 rounded text-sm font-semibold mr-2"
                    style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
                    onClick={exportPackingListSummary}>{t.plrExportBtn}</button>
                  <button className="px-3 py-1.5 rounded text-sm font-semibold mr-2"
                    style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                    onClick={savePackingListRows}>{t.plrSaveBtn}</button>
                  <button className="text-sm font-semibold" style={{ color: colors.inkFaint }}
                    onClick={() => {
                      if (plrRows.length && !window.confirm(t.plrClearConfirm(plrRows.length))) return;
                      setPlrRows([]); setPlrNotes([]); setPlrSaved("");
                      storageSet("packingListReaderRows", "").catch(() => {});
                    }}>{t.plrClearBtn}</button>
                </>
              )}
              {plrRows.length > 0 && (
                <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>
                  {t.plrCount(plrRows.length, new Set(plrRows.map((r) => r["File Name"])).size,
                    plrRows.filter((r) => r.__conflict || r.Check).length)}
                  {plrSaved ? ` \u00b7 ${plrSaved}` : ""}
                </div>
              )}
            </div>
            {plrNotes.map((n, i) => (
              <div key={i} className="rounded px-3 py-2 text-sm" style={{ background: colors.redSoft, color: colors.red }}>{n}</div>
            ))}
            {plrRows.length > 0 && (
              <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}`, background: colors.surface }}>
                <table className="w-full text-sm">
                  <thead><tr style={{ background: colors.surfaceDim }}>
                    {PL_SUMMARY_COLUMNS.map((h) => (
                      <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider whitespace-nowrap" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {plrRows.map((r, i) => (
                      <tr key={i} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        {PL_SUMMARY_COLUMNS.map((c) => {
                          // File Name says where the row came from and Check is worked out
                          // rather than typed, so those two stay read-only. Everything else
                          // is editable: a scan puts the contract number in the client field
                          // often enough that correcting it here, before the download, beats
                          // correcting it in Excel afterwards and then again next time.
                          const readOnly = c === "File Name" || c === "Check";
                          const num = ["PKGS", "KGS", "CBM"].includes(c);
                          const tone = r.__conflict ? colors.red : (c === "Check" && r.Check ? colors.red : colors.ink);
                          if (readOnly) {
                            return (
                              <td key={c} className="px-3 py-1.5"
                                style={{ color: tone, fontWeight: (r.__conflict || (c === "Check" && r.Check)) ? 600 : undefined, maxWidth: c === "Check" ? 300 : 220 }}>
                                {/* Clamped to three lines. A row whose note runs long used to
                                    grow the row past the height of the screen, which made the
                                    table impossible to scroll; the whole note is on hover. */}
                                <div title={String(r[c] || "")}
                                  style={c === "Check" ? { display: "-webkit-box", WebkitLineClamp: 3, WebkitBoxOrient: "vertical", overflow: "hidden", lineHeight: 1.35 } : undefined}>
                                  {r[c] === "" || r[c] === undefined ? "\u2014" : String(r[c])}
                                </div>
                              </td>
                            );
                          }
                          return (
                            <td key={c} className="px-1 py-1">
                              <input
                                className="w-full rounded"
                                style={{
                                  border: `1px solid ${colors.line}`, background: colors.surface,
                                  padding: "3px 6px", fontSize: 13, color: tone,
                                  fontFamily: num ? FONT_MONO : undefined,
                                  textAlign: num ? "right" : "left",
                                  minWidth: c === "Cases" ? 320 : num ? 70 : 130,
                                }}
                                value={r[c] === undefined ? "" : String(r[c])}
                                title={c === "Cases" ? String(r[c] || "") : undefined}
                                onChange={(e) => {
                                  const v = num ? (e.target.value === "" ? "" : Number(e.target.value) || e.target.value) : e.target.value;
                                  setPlrRows((prev) => withConflicts(prev.map((row, n) => (n === i ? { ...row, [c]: v } : row))));
                                }}
                              />
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {view === "ledger" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-4 flex flex-wrap items-start gap-3" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <div style={{ flex: "1 1 380px" }}>
                <h3 className="text-sm font-bold uppercase tracking-wider mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.ledgerTitle}</h3>
                <p className="text-sm mb-2" style={{ color: colors.inkFaint }}>{t.ledgerDesc}</p>
                <div className="flex flex-wrap items-end gap-3">
                  <Field label={t.ledgerAsOfLabel} colors={colors}>
                    <input type="date" className={inputClass} style={inputStyleFor(colors)}
                      value={ledgerAsOf} onChange={(e) => setLedgerAsOf(e.target.value)} />
                  </Field>
                  <button className="text-xs font-semibold pb-2" style={{ color: colors.amberText }}
                    onClick={() => setLedgerAsOf(todayStr())}>{t.ledgerAsOfToday}</button>
                  <div className="text-sm pb-1.5" style={{ color: colors.ink }}>
                    <span className="font-bold" style={{ fontFamily: FONT_DISPLAY, fontSize: 18 }}>{ledgerTotal}</span>
                    <span className="ml-2" style={{ color: colors.inkFaint }}>{t.ledgerTotalOnHand(ledgerAsOf ? fmt(ledgerAsOf) : "")}</span>
                  </div>
                  {ledgerUndatedTotal > 0 && (
                    <div className="text-sm pb-1.5 px-2 py-1 rounded" style={{ background: colors.redSoft, color: colors.red }}>
                      {t.ledgerUndatedTotal(ledgerUndatedTotal)}
                    </div>
                  )}
                </div>
              </div>
              <button className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
                onClick={exportLedger}>{t.ledgerExportBtn}</button>
            </div>
            {ledgerSites.map((g) => {
              const open = ledgerOpen[g.key] !== undefined ? ledgerOpen[g.key] : ledgerSites.length <= 3;
              return (
                <div key={g.key} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
                  <div className="px-4 py-2 flex flex-wrap items-baseline gap-x-3 cursor-pointer" style={{ background: colors.surfaceDim }}
                    onClick={() => setLedgerOpen((o) => ({ ...o, [g.key]: !open }))}>
                    <span className="text-xs font-semibold" style={{ color: colors.amberText }}>{open ? "\u2212" : "+"}</span>
                    <span className="text-sm font-bold" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{g.label}</span>
                    {g.labelZh && <span className="text-xs" style={{ color: colors.inkFaint }}>{g.labelZh}</span>}
                    <span className="text-xs" style={{ color: colors.inkFaint }}>{g.client}</span>
                    <span className="text-xs ml-auto" style={{ color: colors.ink }}>{t.ledgerSiteLine(g.inTotal, g.outTotal, g.balance)}</span>
                  </div>
                  {(g.later > 0 || g.undated > 0) && (
                    <div className="px-4 py-1.5 text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
                      {t.ledgerSetAside(g.later, g.undated)}
                    </div>
                  )}
                  {open && (
                    <table className="w-full text-sm" style={{ background: colors.surface }}>
                      <thead><tr style={{ background: colors.surfaceDim }}>
                        {[t.ledgerColDate, t.ledgerColJob, t.ledgerColDir, t.ledgerColType, t.siteRefsColRef, t.siteRefsColLots, t.ledgerColIn, t.ledgerColOut, t.ledgerColBalance].map((h, hi) => (
                          <th key={hi} className={`px-3 py-1.5 text-xs font-semibold uppercase tracking-wider ${hi >= 6 ? "text-right" : "text-left"}`} style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {(g.undatedMoves || []).map((m, mi) => (
                          <tr key={`u${mi}`} style={{ borderTop: `1px solid ${colors.surfaceDim}`, background: colors.redSoft, color: colors.red }}>
                            <td className="px-3 py-1.5 whitespace-nowrap">
                              <input type="date" className={inputClass}
                                style={{ ...inputStyleFor(colors), padding: "2px 6px", fontSize: 12, borderColor: colors.red }}
                                value=""
                                title={t.ledgerAddDateHint}
                                onChange={(e) => setMovementDate(m, e.target.value)} />
                            </td>
                            <td className="px-3 py-1.5" style={{ fontFamily: FONT_MONO }}>{m.jobNumber || "\u2014"}</td>
                            <td className="px-3 py-1.5 font-semibold">{m.dir}</td>
                            <td className="px-3 py-1.5 text-xs">{m.type || "\u2014"}</td>
                            <td className="px-3 py-1.5 text-xs">{m.ref || "\u2014"}</td>
                            <td className="px-3 py-1.5 text-xs">{m.lot && m.lot !== m.ref ? m.lot : "\u2014"}</td>
                            <td className="px-3 py-1.5 text-right" style={{ fontFamily: FONT_MONO }}>{m.dir === "IN" ? m.pkgs : ""}</td>
                            <td className="px-3 py-1.5 text-right" style={{ fontFamily: FONT_MONO }}>{m.dir === "OUT" ? m.pkgs : ""}</td>
                            <td className="px-3 py-1.5 text-right text-xs">{t.ledgerNotCounted}</td>
                          </tr>
                        ))}
                        {g.moves.map((m, mi) => (
                          <tr key={mi} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                            <td className="px-3 py-1.5 whitespace-nowrap">{m.date ? fmt(m.date) : "\u2014"}</td>
                            <td className="px-3 py-1.5" style={{ fontFamily: FONT_MONO }}>{m.jobNumber || "\u2014"}</td>
                            <td className="px-3 py-1.5 font-semibold" style={{ color: m.dir === "IN" ? colors.green : colors.red }}>{m.dir}</td>
                            <td className="px-3 py-1.5 text-xs" style={{ color: colors.inkFaint }}>{m.type || "\u2014"}</td>
                            <td className="px-3 py-1.5 text-xs">{m.ref || "\u2014"}</td>
                            <td className="px-3 py-1.5 text-xs" style={{ color: colors.inkFaint }}>{m.lot && m.lot !== m.ref ? m.lot : "\u2014"}</td>
                            <td className="px-3 py-1.5 text-right" style={{ fontFamily: FONT_MONO }}>{m.dir === "IN" ? m.pkgs : ""}</td>
                            <td className="px-3 py-1.5 text-right" style={{ fontFamily: FONT_MONO }}>{m.dir === "OUT" ? m.pkgs : ""}</td>
                            <td className="px-3 py-1.5 text-right font-semibold" style={{ fontFamily: FONT_MONO, color: m.balance < 0 ? colors.red : colors.ink }}>{m.balance}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {view === "checkins" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.checkInsTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.checkInsDesc}</p>
              {overCheckedIn > 0 && (
                <p className="text-sm mt-2 rounded px-2 py-1.5" style={{ background: colors.redSoft, color: colors.red }}>{t.checkInsOverBanner(overCheckedIn)}</p>
              )}
              <div className="mt-3">
                <Field label={t.checkInsSearchLabel} colors={colors}>
                  <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 280 }}
                    placeholder={t.checkInsSearchPlaceholder}
                    value={checkInSearch} onChange={(e) => setCheckInSearch(e.target.value)} />
                </Field>
              </div>
            </div>
            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.storeListTitle}</h3>
              <p className="text-sm mb-2" style={{ color: colors.inkFaint }}>{t.storeListDesc}</p>
              <input ref={storeInputRef} type="file" accept=".xlsx,.xls,.xlsm" className="hidden" onChange={handleStoreListFile} />
              <button className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
                onClick={() => storeInputRef.current && storeInputRef.current.click()}>
                {t.storeListChooseBtn}
              </button>
              {storeMoves && (
                <span className="text-xs ml-3" style={{ color: colors.inkFaint }}>
                  {t.storeListLoaded(storeName, storeMoves.length, new Set(storeMoves.map((m) => m.site)).size)}
                </span>
              )}
            </div>
            {storeRec.map((g) => (
              <div key={g.site} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-2" style={{ background: colors.surfaceDim }}>
                  <span className="text-sm font-bold" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{g.site}</span>
                  <span className="text-xs ml-3" style={{ color: g.leftApp === g.leftSheet ? colors.green : colors.red }}>
                    {t.storeListSiteLine(g.inSheet, g.outSheet, g.leftSheet, g.inApp, g.outApp, g.leftApp)}
                  </span>
                </div>
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead><tr style={{ background: colors.surfaceDim }}>
                    {[t.storeListColJob, t.storeListColDir, t.storeListColDm, t.storeListColSheet, t.storeListColApp, t.storeListColDiff].map((h, hi) => (
                      <th key={hi} className="text-left px-3 py-1.5 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {g.rows.filter((r) => r.diff !== 0).map((r, ri) => (
                      <tr key={ri} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <td className="px-3 py-1.5" style={{ fontFamily: FONT_MONO }}>{r.jobNumber}</td>
                        <td className="px-3 py-1.5">{r.direction}</td>
                        <td className="px-3 py-1.5 text-xs" style={{ color: colors.inkFaint }}>{r.dm || "\u2014"}</td>
                        <td className="px-3 py-1.5" style={{ fontFamily: FONT_MONO }}>{r.pkgs}</td>
                        <td className="px-3 py-1.5" style={{ fontFamily: FONT_MONO }}>{r.app}</td>
                        <td className="px-3 py-1.5 font-semibold" style={{ fontFamily: FONT_MONO, color: colors.red }}>{r.diff > 0 ? `+${r.diff}` : r.diff}</td>
                      </tr>
                    ))}
                    {g.rows.every((r) => r.diff === 0) && (
                      <tr><td colSpan={6} className="px-3 py-2 text-sm" style={{ color: colors.green }}>{t.storeListSiteAgrees(g.rows.length)}</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            ))}
            {checkInFiltered.length === 0 && (
              <div className="rounded-lg p-6 text-sm text-center" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.inkFaint }}>
                {t.noneFoundMsg}
              </div>
            )}
            {checkInFiltered.map((g) => {
              const over = (g.over || 0) > 0;
              // Open when there is something wrong with it, or when a search has narrowed it
              // down. A depot's worth of correct check-ins otherwise buries the few that
              // matter under thousands of rows.
              const open = checkInOpen[g.ref] !== undefined ? checkInOpen[g.ref] : (over || !!checkInSearch.trim());
              return (
                <div key={g.ref} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${over ? colors.red : colors.line}` }}>
                  <div className="px-4 py-2 flex flex-wrap items-baseline gap-x-3 cursor-pointer" style={{ background: over ? colors.redSoft : colors.surfaceDim }}
                    onClick={() => setCheckInOpen((o) => ({ ...o, [g.ref]: !open }))}>
                    <span className="text-xs font-semibold" style={{ color: colors.amberText }}>{open ? "\u2212" : "+"}</span>
                    <span className="text-sm font-bold" style={{ fontFamily: FONT_DISPLAY, color: over ? colors.red : colors.ink }}>{g.ref}</span>
                    <span className="text-xs" style={{ color: over ? colors.red : colors.inkFaint }}>
                      {g.stated == null
                        ? t.checkInsHeldOnly(g.held, g.remaining)
                        : t.checkInsHeldVsStated(g.held, g.stated, g.remaining)}
                    </span>
                    {over && <span className="text-xs font-semibold" style={{ color: colors.red }}>{t.checkInsOverBy(g.over)}</span>}
                    <span className="text-xs ml-auto" style={{ color: colors.inkFaint }}>{t.checkInsRowCount(g.rows.length)}</span>
                  </div>
                  {open && (
                  <table className="w-full text-sm" style={{ background: colors.surface }}>
                    <thead>
                      <tr style={{ background: colors.surfaceDim }}>
                        {[t.checkInsColEntry, t.checkInsColCode, t.checkInsColArrived, t.checkInsColUnits, t.checkInsColLeft, t.checkInsColSource, ""].map((h, hi) => (
                          <th key={hi} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {g.rows.map((r, ri) => (
                        <tr key={`${r.item.id}-${ri}`} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                          <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.item.id}</td>
                          <td className="px-3 py-2">{r.item.unitCode || "\u2014"}</td>
                          <td className="px-3 py-2 whitespace-nowrap">{fmt(r.date)}</td>
                          <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.units}</td>
                          <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.left}</td>
                          <td className="px-3 py-2 text-xs" style={{ color: colors.inkFaint, maxWidth: 240 }}>{r.source || "\u2014"}</td>
                          <td className="px-3 py-2 text-right whitespace-nowrap">
                            {g.rows.length > 1 && (
                              <select
                                className="text-xs mr-3 rounded"
                                style={{ ...inputStyleFor(colors), padding: "2px 4px", fontSize: 11 }}
                                value=""
                                onChange={(e) => {
                                  const keepId = e.target.value;
                                  e.target.value = "";
                                  if (!keepId) return;
                                  const dels = activeDeliveries(r.item).length;
                                  if (!window.confirm(t.checkInsMergeConfirm(r.item.id, keepId, r.units, dels))) return;
                                  mergeCheckInInto(r.item, r.arrival, keepId);
                                }}
                              >
                                <option value="">{t.checkInsMergeLabel}</option>
                                {[...new Set(g.rows.map((x) => x.item.id))].filter((id) => id !== r.item.id).map((id) => (
                                  <option key={id} value={id}>{id}</option>
                                ))}
                              </select>
                            )}
                            <button
                              className="text-xs font-semibold"
                              style={{ color: colors.red }}
                              onClick={() => {
                                const left = (r.item.arrivals || []).filter((a) => !r.arrival || a.id !== r.arrival.id).length;
                                const dels = activeDeliveries(r.item).length;
                                // An entry that has been delivered from is not a clean thing
                                // to undo: taking its arrival away leaves deliveries with no
                                // stock behind them. Said before the fact, not after.
                                if (dels > 0 && left === 0 && !window.confirm(t.checkInsReverseHasDeliveries(r.item.id, dels))) return;
                                if (!window.confirm(t.checkInsReverseConfirm(r.item.id, r.units, left === 0))) return;
                                reverseOneCheckIn(r.item, r.arrival);
                              }}
                            >
                              {t.checkInsReverseBtn}
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {view === "duplicates" && (
          <div className="flex flex-col gap-4">
            {duplicateGroups.length === 0 ? (
              <div className="rounded-lg p-6 text-sm text-center" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.inkFaint }}>
                {t.noneFoundMsg}
              </div>
            ) : (
              duplicateGroups.map((group, gi) => {
                const groupIds = group.map((i) => i.id);
                const sorted = [...group].sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
                // Everything in a group already shares client, project, unit, type, arrival
                // date and invoice number - that is what made them a group - so those tell
                // you nothing about which to keep. What matters is where each one came from
                // and what it is carrying, with the fields that actually differ picked out.
                const differs = (get) => new Set(sorted.map((i) => String(get(i) ?? ""))).size > 1;
                const caseSummary = (i) => {
                  const pkgs = i.packages || [];
                  if (!pkgs.length) return i.packageCount ? t.dupFlatCount(i.packageCount) : "\u2014";
                  const first = pkgs[0].code, last = pkgs[pkgs.length - 1].code;
                  return `${pkgs.length} \u00b7 ${first}${pkgs.length > 1 ? `\u2013${last}` : ""}`;
                };
                const sourceOf = (i) => i.notes || "";
                const cols = [
                  { h: t.colId, get: (i) => i.id, mono: true },
                  { h: t.fJobNumber, get: (i) => i.jobNumber || "\u2014", mono: true },
                  { h: t.fJobRef, get: (i) => i.jobRef || "\u2014" },
                  { h: t.fReference, get: (i) => i.shkNumber || "\u2014" },
                  { h: t.fArrivingType, get: (i) => i.arrivingType || "\u2014" },
                  { h: t.colDepot, get: (i) => depotDisplay(i.depot, lang) },
                  { h: t.dupColCases, get: caseSummary },
                  { h: t.fWeight, get: (i) => i.weightKg || "\u2014", mono: true },
                  { h: t.fVolume, get: (i) => i.volumeCbm || "\u2014", mono: true },
                  { h: t.dupColDeliveries, get: (i) => String(activeDeliveries(i).length) },
                  { h: t.colAddedOn, get: (i) => fmt(i.createdAt) },
                ];
                const jobNumbers = [...new Set(group.map((i) => i.jobNumber).filter(Boolean))];
                return (
                  <div key={gi} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.amber}` }}>
                    <div className="px-4 py-2 flex items-center justify-between flex-wrap gap-2" style={{ background: colors.amberSoft }}>
                      <div>
                        <div className="text-sm font-semibold" style={{ color: colors.amberText }}>
                          {group[0].client} · {group[0].project}{group[0].unitCode ? ` · ${group[0].unitCode}` : ""} — {t.matchingEntries(group.length)}
                        </div>
                        <div className="text-xs" style={{ color: colors.amberText }}>
                          {t.dupSharedLine(fmt(group[0].depotArrivalDate), group[0].itemType || "\u2014", jobNumbers.length ? jobNumbers.join(", ") : "\u2014")}
                        </div>
                      </div>
                      <button className="text-xs font-semibold underline" style={{ color: colors.red }} onClick={() => handleDeleteGroup(groupIds)}>
                        {t.deleteAllBtn(group.length)}
                      </button>
                    </div>
                    <div className="overflow-x-auto">
                    <table className="w-full text-sm" style={{ background: colors.surface }}>
                      <thead>
                        <tr style={{ background: colors.surfaceDim }}>
                          {[...cols.map((c) => c.h), t.colStatus, ""].map((h, hi) => (
                            <th key={hi} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider whitespace-nowrap" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.map((i) => {
                          const source = sourceOf(i);
                          const hasDeliveries = activeDeliveries(i).length > 0;
                          return (
                            <React.Fragment key={i.id}>
                              <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                                {cols.map((c, ci) => {
                                  const diff = differs(c.get);
                                  return (
                                    <td key={ci} className="px-3 py-2 whitespace-nowrap"
                                      style={{
                                        fontFamily: c.mono ? FONT_MONO : undefined,
                                        color: diff ? colors.amberText : colors.ink,
                                        fontWeight: diff ? 600 : 400,
                                      }}>
                                      {c.get(i)}
                                    </td>
                                  );
                                })}
                                <td className="px-3 py-2"><StatusBadge item={i} colors={colors} t={t} /></td>
                                <td className="px-3 py-2 text-right whitespace-nowrap">
                                  <button className="text-xs font-semibold mr-3" style={{ color: colors.green }} onClick={() => handleKeepOne(groupIds, i.id)}>{t.keepDeleteBtn}</button>
                                  <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => handleDelete(i.id)}>{t.deleteBtn}</button>
                                </td>
                              </tr>
                              {(source || hasDeliveries) && (
                                <tr style={{ color: colors.inkFaint }}>
                                  <td colSpan={cols.length + 2} className="px-3 pb-2 text-xs">
                                    {source && <span>{t.dupSourceLabel}: {source}</span>}
                                    {hasDeliveries && (
                                      <span style={{ color: colors.red }}>
                                        {source ? " \u00b7 " : ""}{t.dupHasDeliveriesWarn(activeDeliveries(i).length)}
                                      </span>
                                    )}
                                  </td>
                                </tr>
                              )}
                            </React.Fragment>
                          );
                        })}
                      </tbody>
                    </table>
                    </div>
                    <div className="px-4 py-2 text-xs" style={{ background: colors.surfaceDim, color: colors.inkFaint }}>
                      {t.dupDiffHint}
                    </div>
                  </div>
                );
              })
            )}
          </div>
        )}

        {view === "upload" && (
          <UploadPanel
            onImportRows={handleImportRows} onAddIncoming={handleAddIncoming} existingItems={items}
            directory={directory} setDirectory={setDirectory}
            legacyArchive={legacyArchive} setLegacyArchive={setLegacyArchive}
            items={items} incoming={incoming}
            onLegacyImport={handleLegacyImport}
            onLegacyCheckIn={handleCheckIn} onLegacyCheckInBatch={handleCheckInBatch}
            onLegacyDeliver={handleAddCombinedDelivery} onLegacyEnrich={handleLegacyEnrich} onLegacyReverse={handleLegacyReverse}
            colors={colors} t={t} lang={lang}
          />
        )}

        {view === "incoming" && (
          <IncomingPanel incoming={incoming} setIncoming={setIncoming} items={items} directory={directory} setDirectory={setDirectory} employees={employees} onCheckIn={handleCheckIn} onAddIncoming={handleAddIncoming} colors={colors} t={t} lang={lang} />
        )}

        {view === "billing" && (
          <BillingPanel items={items} invoices={invoices} setInvoices={setInvoices} onDeleteItem={handleDelete} onDeleteItems={handleDeleteMany} authUser={authUser} colors={colors} t={t} lang={lang} />
        )}

        {view === "directory" && (
          <DirectoryPanel directory={directory} setDirectory={setDirectory} employees={employees} setEmployees={setEmployees} freeRules={freeRules} setFreeRules={setFreeRules} cbmRates={cbmRates} setCbmRates={setCbmRates} legacyArchive={legacyArchive} setLegacyArchive={setLegacyArchive} items={items} incoming={incoming} onLegacyImport={handleLegacyImport} onLegacyDeliver={handleAddCombinedDelivery} onLegacyEnrich={handleLegacyEnrich} onLegacyReverse={handleLegacyReverse} onLegacyCheckIn={handleCheckIn} onLegacyCheckInBatch={handleCheckInBatch} colors={colors} t={t} lang={lang} />
        )}

        {view === "joblog" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.jobLogTitle}</h3>
              <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.jobLogDesc}</p>
              <div className="flex flex-wrap gap-3 items-end">
                <Field label={t.searchLabel} colors={colors}>
                  <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }}
                    placeholder={t.jobLogSearchPlaceholder}
                    value={jobLogSearch} onChange={(e) => setJobLogSearch(e.target.value)} />
                </Field>
                <Field label={t.jobLogColType} colors={colors}>
                  <select className={inputClass} style={inputStyleFor(colors)} value={jobLogType} onChange={(e) => setJobLogType(e.target.value)}>
                    <option value="All">{t.statusAll}</option>
                    <option value="Devan">{t.jsDevanType}</option>
                    <option value="CFS">{t.jsCfsType}</option>
                    <option value="Delivery">{t.jsDeliveryType}</option>
                  </select>
                </Field>
                <Field label={t.jobLogColClient} colors={colors}>
                  <select className={inputClass} style={inputStyleFor(colors)} value={jobLogClient} onChange={(e) => setJobLogClient(e.target.value)}>
                    <option value="All">{t.statusAll}</option>
                    {jobLogOptions.clients.map((c) => <option key={c}>{c}</option>)}
                  </select>
                </Field>
                <Field label={t.jobLogColSite} colors={colors}>
                  <select className={inputClass} style={{ ...inputStyleFor(colors), maxWidth: 240 }} value={jobLogSite} onChange={(e) => setJobLogSite(e.target.value)}>
                    <option value="All">{t.statusAll}</option>
                    {jobLogOptions.sites.map((v) => <option key={v}>{v}</option>)}
                  </select>
                </Field>
                <Field label={t.jobLogColRecordedBy} colors={colors}>
                  <select className={inputClass} style={inputStyleFor(colors)} value={jobLogRecordedBy} onChange={(e) => setJobLogRecordedBy(e.target.value)}>
                    <option value="All">{t.statusAll}</option>
                    {jobLogOptions.recordedBy.map((v) => <option key={v}>{v}</option>)}
                  </select>
                </Field>
                <Field label={t.jobLogFromLabel} colors={colors}>
                  <input type="date" className={inputClass} style={inputStyleFor(colors)} value={jobLogFrom} onChange={(e) => setJobLogFrom(e.target.value)} />
                </Field>
                <Field label={t.jobLogToLabel} colors={colors}>
                  <input type="date" className={inputClass} style={inputStyleFor(colors)} value={jobLogTo} onChange={(e) => setJobLogTo(e.target.value)} />
                </Field>
                {jobLogFiltersOn && (
                  <button className="text-xs font-semibold pb-2" style={{ color: colors.amberText }} onClick={clearJobLogFilters}>
                    {t.clearBtn}
                  </button>
                )}
              </div>
              <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>
                {t.jobLogCount(jobLogFiltered.length, jobLog.length, new Set(jobLogFiltered.map((r) => r.jobNumber)).size)}
              </div>
            </div>
            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <table className="w-full text-sm" style={{ background: colors.surface }}>
                <thead>
                  <tr style={{ background: colors.surfaceDim }}>
                    {[t.jobLogColJobNo, t.jobLogColType, t.jobLogColDate, t.jobLogColClient, t.jobLogColSite, t.jobLogColRecordedBy, ""].map((h) => (
                      <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {jobLogFiltered.length === 0 && (
                    <tr><td colSpan={7} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{jobLogFiltersOn ? t.jobLogNoMatchMsg : t.jobLogNoneMsg}</td></tr>
                  )}
                  {jobLogFiltered.map((row, idx) => (
                    <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                      <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{row.jobNumber}</td>
                      <td className="px-3 py-2">
                        <Badge tone={row.type === "Delivery" ? "navy" : row.type === "Devan" ? "amber" : "green"} colors={colors}>
                          {row.type === "Devan" ? t.jsDevanType : row.type === "CFS" ? t.jsCfsType : t.jsDeliveryType}
                        </Badge>
                      </td>
                      <td className="px-3 py-2">{fmt(row.date)}</td>
                      <td className="px-3 py-2">{row.client}</td>
                      <td className="px-3 py-2 max-w-[220px] truncate">{row.site}</td>
                      <td className="px-3 py-2">{row.recordedBy || "—"}</td>
                      <td className="px-3 py-2 text-right">
                        <button className="text-xs font-semibold px-2 py-1 rounded" style={{ background: colors.amber, color: colors.ink }} onClick={() => setPrintJobSheet(row.sheet)}>
                          {t.viewReprintBtn}
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === "cancelledjobs" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.cancelledJobsTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.cancelledJobsDesc}</p>
            </div>
            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <table className="w-full text-sm" style={{ background: colors.surface }}>
                <thead>
                  <tr style={{ background: colors.surfaceDim }}>
                    {[t.jobLogColJobNo, t.jobLogColType, t.jobLogColDate, t.jobLogColClient, t.jobLogColSite, t.jobLogColRecordedBy, ""].map((h) => (
                      <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {cancelledJobs.length === 0 && (
                    <tr><td colSpan={7} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.cancelledJobsNoneMsg}</td></tr>
                  )}
                  {cancelledJobs.map((row, idx) => (
                    <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, opacity: 0.75 }}>
                      <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{row.jobNumber}</td>
                      <td className="px-3 py-2">
                        <Badge tone="grey" colors={colors}>
                          {row.type === "Devan" ? t.jsDevanType : row.type === "CFS" ? t.jsCfsType : t.jsDeliveryType}
                        </Badge>
                      </td>
                      <td className="px-3 py-2">{fmt(row.date)}</td>
                      <td className="px-3 py-2">{row.client}</td>
                      <td className="px-3 py-2 max-w-[220px] truncate">{row.site}</td>
                      <td className="px-3 py-2">{row.recordedBy || "—"}</td>
                      <td className="px-3 py-2 text-right whitespace-nowrap">
                        <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => setPrintJobSheet(row.sheet)}>{t.viewReprintBtn}</button>
                        <button className="text-xs font-semibold mr-3" style={{ color: colors.green }} onClick={row.onRestore}>{t.restoreBtn}</button>
                        <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={row.onPurge}>{t.purgeBtn}</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === "add" && (
          <ItemForm
            initial={editing}
            onSave={handleSave}
            onCancel={() => { setEditing(null); setView("inventory"); }}
            onPrintJobSheet={setPrintJobSheet}
            onMoveCases={handleMoveCases}
            directory={directory}
            employees={employees}
            currentUser={currentUser}
            items={items}
            colors={colors}
            t={t}
            lang={lang}
          />
        )}
      </div>

      {printJobSheet && (
        <JobSheetPrint sheet={printJobSheet} onClose={() => setPrintJobSheet(null)} directory={directory} colors={colors} t={t} lang={lang} />
      )}
      {changePasswordOpen && (
        <ChangePasswordModal name={authUser} onClose={() => setChangePasswordOpen(false)} colors={colors} t={t} />
      )}
    </div>
  );
}

