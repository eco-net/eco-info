Opbygning og redigering af Tema arkivet

index.asp viser afholdte temaer, angivet i rsTema som WHERE id=temaid or id=temaid
linker til tema'id'.asp?id
dvs. der skal oprettes en kopi af tema'id'.asp, hvor 'id' rettes:
eks. tema15.asp

tema'id'.asp viser forsiden med kun tema-leder og tema-sektioner
tema-leder er en inklude-fil: inc_leader_'id'.txt, som skal kopieres fra log/ei/home/tema til denne mappe
tema-sektioner er en inklude-fil: inc_sections_'id'.txt, som skal kopieres fra log/ei/home/tema til denne mappe
id rettes i inkluden:
<--#include virtual="/tema_arkiv/inc_leader_14.txt" -->
og i
<--#include virtual="/tema_arkiv/inc_sections_14.txt" -->


tema-leder (inc_leader.txt) linker til temasider.asp
linken rettes til temaside.asp?id='temaid'

så skulle det køre.


