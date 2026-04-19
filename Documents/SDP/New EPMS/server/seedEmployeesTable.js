const bcrypt = require('bcryptjs');
const { User } = require('./models');

const rawData = `1	Dinitha	Active	Jan 1, 2013	N/A	N/A		Mar 5, 1996	1996								13.3
2	Kavindu	Jul 23, 2023	Apr 1, 2018	Aug 1, 2021	Feb 16, 2023											3.8
3	Lanka	Aug 1, 2020	Apr 1, 2018	N/A	N/A											2.3
4	Thiwanka	Nov 15, 2021	Apr 1, 2019	Feb 1, 2021	Aug 16, 2021											2.1
5	Prathibhani	Aug 1, 2021	Jul 1, 2020	N/A	N/A					 						1.1
6	Akash	Apr 1, 2021	Jul 1, 2020	N/A	N/A											0.8
7	Aruni	Active	Sep 1, 2020	N/A	N/A		Jul 11, 1996	1996								5.6
8	Dias	Active	Dec 11, 2020	N/A	N/A	R.D. Chamika Dimash Chandana	May 1, 2001	2001	200112202280	769143965	0779689065 (Amma)	chamikadimash@gmaiil.com	Kudakubukgollawa, Poonawa.	https://drive.google.com/file/d/1z4ZvwnTZnbmp6ai-2VQ1C84Jc4BywYkZ/view?usp=drive_link	https://drive.google.com/file/d/1SSWo0pe4vtT6S6dd1FqzfWdIctVrzSoX/view?usp=drive_link	5.4
9	Isuru	Apr 1, 2021	Dec 1, 2020	N/A	N/A											0.3
10	Pasindu	Jan 1, 2021	Dec 1, 2020	N/A	N/A											0.1
11	Chathumini	Active	Apr 2, 2021	Feb 23, 2022	Jan 25, 2024	R. M. Kawya Chathumini Munasinghe	Apr 20, 2004	2004	200461104588	719666292	0715500175 (Amma)	Kawya.chathu20@gmail.com	Near the water tank, New mannar road, Medawachchiya	https://drive.google.com/file/d/19W8hvDl9NdZz-gAuKS6EJDqMvjV1AUz-/view?usp=drive_link	https://drive.google.com/file/d/11ddl43-cfQB7mEvo5b5D07eHx7lpTsD7/view?usp=drive_link	3.1
12	Nilan	Active	Aug 16, 2021	N/A	N/A	G.P.G Nilan Manoj Jayarathna	Dec 29, 1996	1996	963641583V	779359533	0760865046(Amma)	nilanm2016@gmail.com	No.31 Mudungodha, Hurigaswewa.	https://drive.google.com/file/d/19zrAz3tiYyX_8FYLWM9dzGd0cC1YWyNN/view?usp=drive_link	https://drive.google.com/file/d/16BoasmsWuVcx496rXUjJ6Y5wqFVMxEPm/view?usp=drive_link	4.7
13	Mahesh	Active	Aug 17, 2021	N/A	N/A	Mahesh Prabath Godage	Aug 21, 1993	1993	932342090V	772175509	0741539784 (Malli)	maheshgodage1993@gmail.com	Nabadagaswewa , Rambawa	https://drive.google.com/file/d/1mjemLDpUwaWzcwtkhhQwL9lXLXsl9IvL/view?usp=drive_link	https://drive.google.com/file/d/1fSBvcm18TQRN4utaZ799ZgQz0KNeZDC-/view?usp=drive_link	4.7
14	Thusara	Aug 26, 2021	Aug 18, 2021	N/A	N/A											0.0
15	Buddika	Sep 1, 2021	Aug 1, 2021	N/A	N/A											0.1
16	Nevindi	Active	Feb 20, 2022	Sep 1, 2022	Feb 3, 2023		Aug 8, 2002	2002			0715500175 (Amma)					3.7
17	Chathurya	Apr 1, 2022	Apr 1, 2022	N/A	N/A											0.0
18	Pradeep	May 30, 2025	Jun 20, 2022	Oct 23, 2022	Feb 20, 2024	Gange Gedara Pradeep Asanka Dharmasena	Jul 11, 1993	1993	931931563V	761309861	0711452771 (Wife)	pradeep2asanka@gmail.com	Manaram Plaugaswewa, Galkulama, Anuradhapura	https://drive.google.com/file/d/1HfOnszsgb5B-31ommmOUGSGrUVv86O0E/view?usp=drive_link	https://drive.google.com/file/d/1L1Nr1isY6rwPX-eWIAjro7oyJ3O7etO1/view?usp=drive_link	1.6
19	Desha	Mar 25, 2024	Feb 2, 2024	N/A	N/A	Herath Mudiyanselage Desha Nirmani Herath	Jun 21, 1995		956732247V	719073617	0718396856 (Amma)	d.nirmanih@gmail.com	No 39, Jayasiripura Farm, Jayasiripura, Anuradhapura			0.1
20	Razeena	May 13, 2024	Feb 5, 2024	N/A	N/A	Razeena Marso	Jul 7, 2003		200368910574	750269555	0716734610 (sisters)	razeemarso130@gmail.com	53/5TH LANE ANURADHAPURA ROAD PUTTALAM	https://drive.google.com/file/d/1VGiPGG2sqSZF-bU8aUwf4oHckEGeFdGn/view?usp=drive_link	https://drive.google.com/file/d/1_wVbYMj1Zr7rMBgJWlTD36SmgX2U2U8z/view?usp=drive_link	0.3
21	Audarya	Oct 10, 2024	Mar 4, 2024	N/A	N/A	Herath Mudiyanselage Audarya Samadhi Herath	May 25, 1992		921460309V	0777492250, 0777338500	0372253244 (Home)	audarya@live.com	Dispensery, Nelumpathwewa Rd, Ambanpola	https://drive.google.com/file/d/1s6KVGh5i1OVfUIq8sRFBVB0tTjeTMs8z/view?usp=drive_link	https://drive.google.com/file/d/1O32lkeBCF1sPykbK9ZBZlLfEBNgojkLP/view?usp=drive_link	0.6
22	Thilina	Apr 26, 2024	Mar 27, 2024	N/A	N/A	Thilina Saubhagya Godallage	Dec 20, 2000		200035504193	717724876	0712744172 (Amma)	thilinagodallagecentral@gmail.com	Hiruna, yodha ela road,5/ela, Srawasthipura, Anuradhapura	https://drive.google.com/file/d/1QhheBhtES33vQRcVZqb-7V_07gVh2cQz/view?usp=drive_link	https://drive.google.com/file/d/1vSaY0fJJ368IF0JmzPfdeC1YenhFfcFk/view?usp=drive_link	0.1
23	Praveen	Active	Apr 5, 2024	N/A	N/A	Praveen Madhushankha Weerasinghe	Aug 30, 1996	1996	962431372V	773496874	0765869782 (Wife)	praveenw96@gmail.com	21/154A1, Vijaya Mawatha, Isurupura, Anuradhapura	https://drive.google.com/file/d/11te3unJaUNWrusXnmKyh8RaWIqPBJziT/view?usp=drive_link	https://drive.google.com/file/d/11te3unJaUNWrusXnmKyh8RaWIqPBJziT/view?usp=drive_link	2.0
24	Hasith	Nov 2, 2024	Sep 24, 2024	N/A	N/A	Hasith Sankhaja Abeyweera	Apr 13, 1998		981042743V	771266608	076 4 191 582 - Asha Pradeepika(Mother)	hasithabeweera@gmail.com	No: 80/4 Walawwatta 2nd Step,Galahitiyawa,Madampe.			0.1
25	Malshi	Oct 24, 2025	Dec 30, 2024	N/A	N/A	Nalaka Priyanthage Malshi Parindya	Jul 10, 2004	2004	200469203532	761404907	718989404(mother)	malshiparindya20@gmail.com	Daham Mawatha,Postal Village,Medawachchiya.			0.8
26	Lakmal	Mar 11, 2025	Jan 19, 2025	N/A	N/A	Wickramasingha senanayake appuhamillage Lakmal bandara senanayake	Mar 23, 1989		890830137V	719915756	374908004	lakmalbandarasenanayake@gmail.com	mary land,bohingamuwa,kuliyapitiya			0.1
27	Imodya	Sep 2, 2025	Feb 10, 2025	N/A	N/A	Basnayake Mudiyanselage Imodya Basnayake	Oct 24, 1999	1999	997981278V	710966438	718025661 - J.P.M Janakanthi ( Mother)	imodya.99@gmail.com	No.6, Samagi Mawatha, K.W.A Junction, wanniyankulama, Anuradhapura			0.6
28	Madhushanka	Active	May 4, 2025	N/A	N/A	Aththanayaka Mudiyanselage Sudath Madhushanka Aththanayaka	Mar 5, 1996	1996	960651960V	701367490	767673916		No 115, Morawakakanda, Nochchiyagama			1.0
29	Pamudini	May 20, 2025	May 12, 2025	N/A	N/A	Subhasinghe Mudiyanselage Pamudini Nimeshika Karunarathne 	Oct 14, 2000	2000		742204445		pamudinikarunarathne58@gmail.com				0.0
30	Shirantha	Active	Jun 5, 2025	N/A	N/A	Halawathage Shirantha Meril Damion Perera	Sep 28, 1998	1998	982720648V	0714886107	0714456162 - Mala Bopitiya (mother)	shiranthaperera36@gmail.com				0.9
31	Nick	Active	Jun 6, 2025	N/A	N/A	Gamage Nilan Nadeesha	Mar 16, 2001	2001	200107603890	0772595531	0742595531 - Kusum - (Mother)	nilannadeesha.ceo@gmail.com	154/B, Pathegama, Rammala, Warapitiya.			0.9
32	Lasitha	Jun 20, 2025	Jun 6, 2025	N/A	N/A	Baminahannadige Lasitha Madhusanka Peiris	Dec 30, 1996	1996	963654324V	0712360747	0761747332 - Sarath Peiris (Father)	lasimadhusanka@gmail.com	No.25, Pubudupura, Anuradhapura			0.0
33	Chanaka	Jul 25, 2025	Jun 11, 2025	N/A	N/A	Y.M Chanaka Madusanka Abeyrathna	Jul 5, 1991	1991	911871688v	0714904992	0706808083 - Gayani	ymcmabeyrathna@gmail.com	No.18 , pansalawatta Rd. Nattarampotha.			0.1
34	Sahan	Nov 17, 2025	Sep 15, 2025	N/A	N/A	W.D. Sahan Rashmika	Jun 10, 2002		200216203089	+94765412462	0252224022 - Home	rashmikasahan228@gmail.com	30/69 baduwatta avenue, ramakale, walawwatta,anuradhapura			0.2
35	Thanuja	Oct 7, 2025	Sep 24, 2025	N/A	N/A	Herath Mudiyanselage Thanuja Dhananjaya Ranasinghe	Oct 19, 2000		200029300736	0701530191	0714093132 - C Ranasinghe - Father	thanujadranasinghe@gmail.com	No 457/10, 3rd lane, Nidahas mwt, Pullayar Junction, Anuradhapura.			0.0
36	Nilupul		Jan 16, 2026			M.M.D. Nilupul Nishada Gunawardhana 	Aug 6, 1998	1998	982191181v	0789583978 / 0705705837	0727763260 - Moksha Randuli (Sister)	nilupulnishada9@gmail.com	No.99, U.C.Quarters, New Town, Anuradhapura 			\n`;

const seedParsedEmployees = async () => {
    try {
        const lines = rawData.split('\n').filter(l => l.trim().length > 0);
        const salt = await bcrypt.genSalt(10);
        const hashedPassword = await bcrypt.hash('password123', salt);

        let addedCount = 0;

        for (const line of lines) {
            const cols = line.split('\t');
            const name = cols[1]?.trim() || 'Unknown Employee';
            let emailMatch = line.match(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b/);
            let email = emailMatch ? emailMatch[0].toLowerCase() : `${name.toLowerCase().replace(/[^a-z]/g, '')}${cols[0]}@techznap.com`;

            let uniqueUsername = name;
            let existingUser = await User.findOne({ where: { username: uniqueUsername } });
            if (existingUser && existingUser.email !== email) {
                uniqueUsername = `${name}_${cols[0]}`;
            }

            await User.findOrCreate({
                where: { email: email },
                defaults: {
                    username: uniqueUsername,
                    role: 'Employee',
                    password: hashedPassword,
                    position: 'Employee'
                }
            });
            addedCount++;
        }

        console.log(`Successfully processed and added ${addedCount} employees!`);
        process.exit(0);
    } catch (err) {
        console.error('Error seeding employees:', err);
        process.exit(1);
    }
}

seedParsedEmployees();
