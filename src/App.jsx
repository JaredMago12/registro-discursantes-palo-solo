import React, { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import { LogOut, LogIn, Download, Trash2, ChevronDown, MapPin, Phone, Clock3, AlertCircle, Search, X } from 'lucide-react';
import * as XLSX from 'xlsx';

const STORAGE_KEY = 'palo-solo:registro-discursantes';

const ROSTER = [
  { turno: 1, nombre: 'Eduardo López', rol: 'Anciano', tel: '55 1432 7312',
    preparados: [19, 71, 145, 186] },
  { turno: 2, nombre: 'Daniel García', rol: 'Anciano', tel: '55 2860 7385',
    preparados: [51, 99, 145, 193] },
  { turno: 3, nombre: 'Félix Hernández', rol: 'Anciano', tel: '55 3521 2934',
    preparados: [151] },
  { turno: 4, nombre: 'Jared Sánchez', rol: 'Anciano', tel: '56 1785 5084',
    preparados: [142, 3, 46, 32, 53] },
  { turno: 5, nombre: 'Juan Carlos Morales', rol: 'Ministerial', tel: '55 6067 1525',
    preparados: [103, 2] },
  { turno: 6, nombre: 'Ithiel Tabaco', rol: 'Ministerial', tel: '55 6669 6666',
    preparados: [32, 113, 116] },
  { turno: 7, nombre: 'Javier Santos', rol: 'Ministerial (Coordinador)', tel: '55 8469 2565',
    preparados: [78, 93] },
];

const DISCURSOS_OFICIALES = [
  {n:1,t:'¿Conoce bien a Dios?',c:'Fe/Espiritualidad',d:true},
  {n:2,t:'¿Sobreviviremos a los últimos días?',c:'Últimos días/Juicio de Dios',d:true},
  {n:3,t:'Avancemos con la organización unida de Jehová',c:'Religión/Adoración',d:true},
  {n:4,t:'El mundo que nos rodea prueba que Dios existe',c:'Biblia/Dios',d:true},
  {n:5,t:'Ayuda práctica para las familias',c:'Familia/Jóvenes',d:true},
  {n:6,t:'Qué aprendemos del diluvio universal',c:'Últimos días/Juicio de Dios',d:true},
  {n:7,t:'Imitemos al "Padre de tiernas misericordias"',c:'Normas y cualidades cristianas',d:true},
  {n:8,t:'Vivamos para hacer la voluntad de Dios, no la nuestra',c:'Religión/Adoración',d:true},
  {n:9,t:'Escuchemos y pongamos en práctica la Palabra de Dios',c:'Fe/Espiritualidad',d:true},
  {n:10,t:'Seamos honrados en todo',c:'Normas y cualidades cristianas',d:true},
  {n:11,t:'Imitemos a Jesús, y no seamos parte del mundo',c:'Mundo, no ser parte del',d:true},
  {n:12,t:'A Dios le importa cómo vemos la autoridad',c:'Normas y cualidades cristianas',d:true},
  {n:13,t:'Cómo ve Dios el sexo y el matrimonio',c:'Familia/Jóvenes',d:true},
  {n:14,t:'Un pueblo limpio da gloria a Jehová',c:'Normas y cualidades cristianas',d:true},
  {n:15,t:'"Hagamos el bien a todos"',c:'Normas y cualidades cristianas',d:true},
  {n:16,t:'Fortalezcamos nuestra amistad con Dios',c:'Fe/Espiritualidad',d:true},
  {n:17,t:'Demos gloria a Dios con todo lo que tenemos',c:'Evangelización/Ministerio',d:true},
  {n:18,t:'Haga de Jehová su "fortaleza"',c:'Fe/Espiritualidad',d:true},
  {n:19,t:'Cómo puede usted conocer su futuro',c:'Reino/Paraíso',d:true},
  {n:20,t:'¿Ha llegado el tiempo para que Dios gobierne el mundo?',c:'Últimos días/Juicio de Dios',d:true},
  {n:21,t:'Valoremos nuestro lugar en el Reino de Dios',c:'Reino/Paraíso',d:true},
  {n:22,t:'¿Aprovecha usted todas las ayudas espirituales que Jehová nos da?',c:'Fe/Espiritualidad',d:true},
  {n:23,t:'La vida tiene propósito',c:'Reino/Paraíso',d:true},
  {n:24,t:'¿Ha encontrado usted "una perla muy valiosa"?',c:'Reino/Paraíso',d:true},
  {n:25,t:'Luchemos contra el espíritu del mundo',c:'Mundo, no ser parte del',d:true},
  {n:26,t:'¿Le importamos a Dios?',c:'Biblia/Dios',d:true},
  {n:27,t:'Cómo iniciar bien el matrimonio',c:'Familia/Jóvenes',d:true},
  {n:28,t:'Muestre respeto y amor en su matrimonio',c:'Familia/Jóvenes',d:true},
  {n:29,t:'Las responsabilidades y las recompensas que tienen los padres',c:'Familia/Jóvenes',d:true},
  {n:30,t:'Cómo mejorar la comunicación en la familia',c:'Familia/Jóvenes',d:true},
  {n:31,t:'¿Estamos al tanto de nuestras necesidades espirituales?',c:'Fe/Espiritualidad',d:true},
  {n:32,t:'Cómo enfrentarse a las inquietudes de la vida',c:'Pruebas/Problemas',d:true},
  {n:33,t:'¿Habrá algún día justicia para todos?',c:'Mundo, no ser parte del',d:true},
  {n:34,t:'¿Tendrá usted la marca para sobrevivir?',c:'Últimos días/Juicio de Dios',d:true},
  {n:35,t:'¿Se puede vivir para siempre? ¿Lo logrará usted?',c:'Reino/Paraíso',d:true},
  {n:36,t:'¿Es esta vida todo lo que podemos esperar?',c:'Religión/Adoración',d:true},
  {n:37,t:'¿Por qué andar en el camino de Dios?',c:'Biblia/Dios',d:true},
  {n:38,t:'¿Cómo puede usted sobrevivir al fin del mundo?',c:'Últimos días/Juicio de Dios',d:true},
  {n:39,t:'¿En qué sentido es Jesucristo el vencedor del mundo?',c:'Mundo, no ser parte del',d:true},
  {n:40,t:'¿Qué sucederá en el futuro cercano?',c:'Últimos días/Juicio de Dios',d:true},
  {n:41,t:'"Estense quietos y vean cómo los salva Jehová"',c:'Últimos días/Juicio de Dios',d:true},
  {n:42,t:'¿Puede el amor vencer al odio?',c:'Normas y cualidades cristianas',d:true},
  {n:43,t:'Lo que Dios espera de nosotros siempre nos beneficia',c:'Religión/Adoración',d:true},
  {n:44,t:'¿Cómo le benefician a usted las enseñanzas de Jesús?',c:'Fe/Espiritualidad',d:true},
  {n:45,t:'Sigamos el camino que lleva a la vida',c:'Religión/Adoración',d:true},
  {n:46,t:'Mantengamos nuestra confianza fuerte hasta el fin',c:'Fe/Espiritualidad',d:true},
  {n:47,t:'"Tengan fe en las buenas noticias"',c:'Reino/Paraíso',d:true},
  {n:48,t:'Cómo ser leales ante las pruebas',c:'Normas y cualidades cristianas',d:true},
  {n:49,t:'¿Viviremos algún día en una Tierra limpia?',c:'Reino/Paraíso',d:true},
  {n:50,t:'Cómo tomar buenas decisiones',c:'Pruebas/Problemas',d:true},
  {n:51,t:'¿Está la verdad transformando su vida?',c:'Mundo, no ser parte del',d:true},
  {n:52,t:'¿Quién es su Dios?',c:'Religión/Adoración',d:true},
  {n:53,t:'¿Piensa usted igual que Dios?',c:'Mundo, no ser parte del',d:true},
  {n:54,t:'Tenga fe en Dios y sus promesas',c:'Biblia/Dios',d:true},
  {n:55,t:'¿Cómo puede ganarse una buena reputación ante Dios?',c:'Religión/Adoración',d:true},
  {n:56,t:'¿En qué líder podemos confiar?',c:'Religión/Adoración',d:true},
  {n:57,t:'Aguantemos la persecución',c:'Pruebas/Problemas',d:true},
  {n:58,t:'¿Quiénes son los verdaderos discípulos de Cristo?',c:'Religión/Adoración',d:true},
  {n:59,t:'(No usar)',c:'Mundo, no ser parte del',d:false},
  {n:60,t:'¿Cuál es su propósito en la vida?',c:'Fe/Espiritualidad',d:true},
  {n:61,t:'¿En las promesas de quién confía usted?',c:'Reino/Paraíso',d:true},
  {n:62,t:'¿Dónde encontrará una esperanza segura?',c:'Reino/Paraíso',d:true},
  {n:63,t:'¿Es posible encontrar la verdad?',c:'Biblia/Dios',d:true},
  {n:64,t:'¿Amamos los placeres en vez de a Dios?',c:'Mundo, no ser parte del',d:true},
  {n:65,t:'¿Cómo podemos ser pacíficos en un mundo violento?',c:'Pruebas/Problemas',d:true},
  {n:66,t:'¿Será usted un buen trabajador en la cosecha?',c:'Evangelización/Ministerio',d:true},
  {n:67,t:'Medite en la Palabra de Jehová y en su creación',c:'Fe/Espiritualidad',d:true},
  {n:68,t:'"Sigan [...] perdonándose con generosidad"',c:'Normas y cualidades cristianas',d:true},
  {n:69,t:'¿Por qué es importante mostrar amor desinteresado?',c:'Normas y cualidades cristianas',d:true},
  {n:70,t:'¿Por qué se merece Dios toda nuestra confianza?',c:'Biblia/Dios',d:true},
  {n:71,t:'Ahora es el momento de estar despiertos',c:'Fe/Espiritualidad',d:true},
  {n:72,t:'El amor identifica a la religión verdadera',c:'Normas y cualidades cristianas',d:true},
  {n:73,t:'Consigamos que nuestro corazón sea sabio',c:'Pruebas/Problemas',d:true},
  {n:74,t:'Jehová está pendiente de nosotros',c:'Fe/Espiritualidad',d:true},
  {n:75,t:'Demuestre que apoya el gobierno de Dios',c:'Normas y cualidades cristianas',d:true},
  {n:76,t:'¿Pueden los principios bíblicos ayudarnos a afrontar los problemas de nuestro tiempo?',c:'Biblia/Dios',d:true},
  {n:77,t:'"Sean siempre hospitalarios"',c:'Normas y cualidades cristianas',d:true},
  {n:78,t:'Sirva a Jehová con alegría',c:'Normas y cualidades cristianas',d:true},
  {n:79,t:'¿A quién prefiere como amigo?',c:'Mundo, no ser parte del',d:true},
  {n:80,t:'¿Confía usted en la ciencia o en la Biblia?',c:'Biblia/Dios',d:true},
  {n:81,t:'¿Quiénes están preparados para hacer discípulos?',c:'Evangelización/Ministerio',d:true},
  {n:82,t:'(No usar)',c:'Religión/Adoración',d:false},
  {n:83,t:'¿Tienen que obedecer los cristianos los Diez Mandamientos?',c:'Religión/Adoración',d:true},
  {n:84,t:'¿Escapará usted de lo que le espera a este mundo?',c:'Últimos días/Juicio de Dios',d:true},
  {n:85,t:'Buenas nuevas en un mundo violento',c:'Reino/Paraíso',d:true},
  {n:86,t:'Cómo lograr que Dios escuche nuestras oraciones',c:'Religión/Adoración',d:true},
  {n:87,t:'¿Qué relación tiene usted con Dios?',c:'Fe/Espiritualidad',d:true},
  {n:88,t:'Por qué vivir de acuerdo con las normas de la Biblia',c:'Biblia/Dios',d:true},
  {n:89,t:'¡Venga a beber el agua de la verdad!',c:'Religión/Adoración',d:true},
  {n:90,t:'Esfuércese por conseguir la vida que realmente es vida',c:'Reino/Paraíso',d:true},
  {n:91,t:'La presencia y la gobernación del Mesías',c:'Reino/Paraíso',d:true},
  {n:92,t:'El papel de la religión en los asuntos del mundo',c:'Religión/Adoración',d:true},
  {n:93,t:'¿Cuándo se acabarán los desastres naturales?',c:'Pruebas/Problemas',d:true},
  {n:94,t:'La religión verdadera satisface las necesidades de la sociedad humana',c:'Religión/Adoración',d:true},
  {n:95,t:'¡No caiga en la trampa del ocultismo!',c:'Religión/Adoración',d:true},
  {n:96,t:'¿Cuál es el futuro de la religión?',c:'Religión/Adoración',d:true},
  {n:97,t:'Permanezcamos sin culpa en medio de una generación torcida',c:'Mundo, no ser parte del',d:true},
  {n:98,t:'"La escena de este mundo está cambiando"',c:'Últimos días/Juicio de Dios',d:true},
  {n:99,t:'Por qué se puede confiar en la Biblia',c:'Biblia/Dios',d:true},
  {n:100,t:'Haga amistades que duren para siempre',c:'Normas y cualidades cristianas',d:true},
  {n:101,t:'Jehová, el "Gran Creador"',c:'Biblia/Dios',d:true},
  {n:102,t:'Prestémosle atención a "la palabra profética"',c:'Últimos días/Juicio de Dios',d:true},
  {n:103,t:'Cómo ser verdaderamente felices',c:'Normas y cualidades cristianas',d:true},
  {n:104,t:'Padres, ¿están construyendo con materiales resistentes al fuego?',c:'Familia/Jóvenes',d:true},
  {n:105,t:'Cómo obtener consuelo en toda tribulación',c:'Pruebas/Problemas',d:true},
  {n:106,t:'Dios dará su merecido a los que arruinan la tierra',c:'Últimos días/Juicio de Dios',d:true},
  {n:107,t:'Eduquemos nuestra conciencia, y nos irá bien',c:'Mundo, no ser parte del',d:true},
  {n:108,t:'Mire al futuro sin miedo',c:'Pruebas/Problemas',d:true},
  {n:109,t:'El Reino de Dios está cerca',c:'Reino/Paraíso',d:true},
  {n:110,t:'La familia feliz es la que pone a Dios en primer lugar',c:'Familia/Jóvenes',d:true},
  {n:111,t:'La humanidad recuperará la salud por completo',c:'Reino/Paraíso',d:true},
  {n:112,t:'Mostremos amor en un mundo egoísta',c:'Normas y cualidades cristianas',d:true},
  {n:113,t:'Joven, ¿cómo puedes ser feliz y tener éxito?',c:'Familia/Jóvenes',d:true},
  {n:114,t:'Aprecie las maravillas de la creación de Dios',c:'Biblia/Dios',d:true},
  {n:115,t:'Protéjase de las astutas trampas de Satanás',c:'Mundo, no ser parte del',d:true},
  {n:116,t:'Sea sabio al elegir sus compañías',c:'Mundo, no ser parte del',d:true},
  {n:117,t:'Cómo vencer el mal con el bien',c:'Pruebas/Problemas',d:true},
  {n:118,t:'Veamos a los jóvenes como los ve Jehová',c:'Familia/Jóvenes',d:true},
  {n:119,t:'Por qué es provechoso que los cristianos se mantengan separados del mundo',c:'Mundo, no ser parte del',d:true},
  {n:120,t:'Razones para someterse hoy al gobierno de Dios',c:'Reino/Paraíso',d:true},
  {n:121,t:'Una hermandad mundial sobrevivirá a la mayor calamidad de la historia',c:'Últimos días/Juicio de Dios',d:true},
  {n:122,t:'(No usar)',c:'Reino/Paraíso',d:false},
  {n:123,t:'(No usar)',c:'Mundo, no ser parte del',d:false},
  {n:124,t:'Razones para confiar en el origen divino de la Biblia',c:'Biblia/Dios',d:true},
  {n:125,t:'Por qué necesita la humanidad un rescate',c:'Religión/Adoración',d:true},
  {n:126,t:'¿Quién puede ser salvo?',c:'Religión/Adoración',d:true},
  {n:127,t:'¿Qué nos sucede cuando morimos?',c:'Religión/Adoración',d:true},
  {n:128,t:'¿Es el infierno realmente un lugar de tormento?',c:'Religión/Adoración',d:true},
  {n:129,t:'¿Enseña la Biblia que Dios es una trinidad?',c:'Religión/Adoración',d:true},
  {n:130,t:'La Tierra existirá para siempre',c:'Reino/Paraíso',d:true},
  {n:131,t:'Pongámonos en contra del Diablo',c:'Mundo, no ser parte del',d:true},
  {n:132,t:'La resurrección: una victoria sobre la muerte',c:'Reino/Paraíso',d:true},
  {n:133,t:'¿Importa lo que creemos sobre el origen de los seres humanos?',c:'Biblia/Dios',d:true},
  {n:134,t:'¿Deben los cristianos observar el sábado?',c:'Religión/Adoración',d:true},
  {n:135,t:'La santidad de la vida y la sangre',c:'Religión/Adoración',d:true},
  {n:136,t:'¿Aprueba Dios las imágenes en la adoración?',c:'Religión/Adoración',d:true},
  {n:137,t:'¿De verdad tuvieron lugar los milagros de la Biblia?',c:'Biblia/Dios',d:true},
  {n:138,t:'Vivamos con buen juicio en un mundo perverso',c:'Mundo, no ser parte del',d:true},
  {n:139,t:'Sabiduría piadosa en un mundo de orientación científica',c:'Biblia/Dios',d:true},
  {n:140,t:'¿Quién es Jesucristo en realidad?',c:'Religión/Adoración',d:true},
  {n:141,t:'¿Cuándo dejará de gemir la creación humana?',c:'Pruebas/Problemas',d:true},
  {n:142,t:'Por qué refugiarse en Jehová',c:'Fe/Espiritualidad',d:true},
  {n:143,t:'Confiemos en el Dios de todo consuelo',c:'Pruebas/Problemas',d:true},
  {n:144,t:'Una congregación leal bajo la dirección de Cristo',c:'Normas y cualidades cristianas',d:true},
  {n:145,t:'¿Quién es como Jehová nuestro Dios?',c:'Biblia/Dios',d:true},
  {n:146,t:'Utilice la educación para alabar a Jehová',c:'Familia/Jóvenes',d:true},
  {n:147,t:'Confiemos en que Jehová tiene el poder para salvarnos',c:'Fe/Espiritualidad',d:true},
  {n:148,t:'¿Ve usted la vida como la ve Dios?',c:'Normas y cualidades cristianas',d:true},
  {n:149,t:'¿Anda usted con Dios?',c:'Fe/Espiritualidad',d:true},
  {n:150,t:'¿Está este mundo condenado a la destrucción?',c:'Últimos días/Juicio de Dios',d:true},
  {n:151,t:'Jehová es "una altura segura" para su pueblo',c:'Fe/Espiritualidad',d:true},
  {n:152,t:'¿Cuándo y por qué vendrá el verdadero Armagedón?',c:'Últimos días/Juicio de Dios',d:true},
  {n:153,t:'Estemos muy pendientes del "impresionante día de Jehová"',c:'Últimos días/Juicio de Dios',d:true},
  {n:154,t:'El gobierno del hombre, pesado en la balanza',c:'Reino/Paraíso',d:true},
  {n:155,t:'¿Ha llegado la hora del juicio de Babilonia?',c:'Religión/Adoración',d:true},
  {n:156,t:'¿Es el Día del Juicio un tiempo de temor, o de esperanza?',c:'Últimos días/Juicio de Dios',d:true},
  {n:157,t:'Cómo adornan los cristianos la enseñanza divina',c:'Normas y cualidades cristianas',d:true},
  {n:158,t:'Seamos valerosos y confiemos en Jehová',c:'Fe/Espiritualidad',d:true},
  {n:159,t:'Cómo encontrar seguridad en un mundo peligroso',c:'Fe/Espiritualidad',d:true},
  {n:160,t:'Protejamos nuestra identidad cristiana',c:'Mundo, no ser parte del',d:true},
  {n:161,t:'¿Por qué sufrió y murió Jesús?',c:'Religión/Adoración',d:true},
  {n:162,t:'Liberados de un mundo de oscuridad',c:'Reino/Paraíso',d:true},
  {n:163,t:'¿Por qué temer al Dios verdadero?',c:'Religión/Adoración',d:true},
  {n:164,t:'¿Sigue Dios ejerciendo la soberanía sobre la Tierra?',c:'Biblia/Dios',d:true},
  {n:165,t:'¿A qué valores concedemos más importancia?',c:'Normas y cualidades cristianas',d:true},
  {n:166,t:'¿Qué es la verdadera fe, y cómo se demuestra?',c:'Fe/Espiritualidad',d:true},
  {n:167,t:'Actuemos con sabiduría en este mundo insensato',c:'Mundo, no ser parte del',d:true},
  {n:168,t:'Cómo sentirse seguro en este mundo turbulento',c:'Fe/Espiritualidad',d:true},
  {n:169,t:'¿Por qué debemos guiarnos por la Biblia?',c:'Biblia/Dios',d:true},
  {n:170,t:'¿Quién es el único que puede gobernar bien a la humanidad?',c:'Reino/Paraíso',d:true},
  {n:171,t:'Usted puede disfrutar de la vida en paz ahora y para siempre',c:'Normas y cualidades cristianas',d:true},
  {n:172,t:'¿Qué posición tenemos ante Dios?',c:'Fe/Espiritualidad',d:true},
  {n:173,t:'¿Le importa a Dios qué religión tengamos?',c:'Religión/Adoración',d:true},
  {n:174,t:'¿Quién será digno de entrar en el nuevo mundo de Dios?',c:'Reino/Paraíso',d:true},
  {n:175,t:'¿Qué pruebas demuestran la autenticidad de la Biblia?',c:'Biblia/Dios',d:true},
  {n:176,t:'¿Cuándo tendremos verdadera paz y seguridad?',c:'Últimos días/Juicio de Dios',d:true},
  {n:177,t:'¿A quién podemos acudir en tiempos de angustia?',c:'Pruebas/Problemas',d:true},
  {n:178,t:'Andemos en el camino de la integridad',c:'Mundo, no ser parte del',d:true},
  {n:179,t:'Rechace las fantasías mundanales; busque las realidades del Reino',c:'Mundo, no ser parte del',d:true},
  {n:180,t:'¿Por qué debe ser real para nosotros la esperanza de la resurrección?',c:'Reino/Paraíso',d:true},
  {n:181,t:'¿Faltará menos de lo que usted cree?',c:'Últimos días/Juicio de Dios',d:true},
  {n:182,t:'¿Qué está haciendo por nosotros el Reino de Dios?',c:'Reino/Paraíso',d:true},
  {n:183,t:'Apartemos la mirada de lo que no sirve para nada',c:'Mundo, no ser parte del',d:true},
  {n:184,t:'¿Es la muerte el final de todo?',c:'Pruebas/Problemas',d:true},
  {n:185,t:'¿Influye la verdad en su vida?',c:'Normas y cualidades cristianas',d:true},
  {n:186,t:'Unidos al feliz pueblo de Dios',c:'Pruebas/Problemas',d:true},
  {n:187,t:'¿Cómo es posible que un Dios de amor permita la maldad?',c:'Biblia/Dios',d:true},
  {n:188,t:'¿Ha puesto usted su confianza en Jehová?',c:'Fe/Espiritualidad',d:true},
  {n:189,t:'Andar con Dios nos beneficia ahora y para siempre',c:'Fe/Espiritualidad',d:true},
  {n:190,t:'Una familia unida y feliz para siempre',c:'Familia/Jóvenes',d:true},
  {n:191,t:'Cómo vencen al mundo la fe y el amor',c:'Mundo, no ser parte del',d:true},
  {n:192,t:'¿Anda usted por el camino que lleva a la vida eterna?',c:'Fe/Espiritualidad',d:true},
  {n:193,t:'Pronto se nos librará de la angustia mundial',c:'Últimos días/Juicio de Dios',d:true},
  {n:194,t:'Cómo nos beneficia la sabiduría divina',c:'Pruebas/Problemas',d:true},
];

const DISCURSOS_POR_NUM = Object.fromEntries(DISCURSOS_OFICIALES.map((d) => [d.n, d]));
const CONGREGACIONES_SUGERIDAS = ['Parque la punta', 'Country', 'Burocrática', 'Palo Solo'];

function normalize(s) {
  return (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

function proximoSabado() {
  const d = new Date();
  const day = d.getDay();
  d.setDate(d.getDate() + ((6 - day + 7) % 7));
  return d.toISOString().slice(0, 10);
}

function fechaLegible(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function nuevoId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function TalkPicker({ value, onChange, placeholder }) {
  const [query, setQuery] = useState(value ? `${value.num} — ${value.titulo}` : '');
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);

  useEffect(() => {
    setQuery(value ? `${value.num} — ${value.titulo}` : '');
  }, [value]);

  const results = useMemo(() => {
    if (!open) return [];
    const qRaw = query.trim();
    const q = normalize(qRaw);
    let list = DISCURSOS_OFICIALES.filter((d) => d.d);
    if (q) {
      list = list.filter((d) => String(d.n).startsWith(qRaw) || normalize(d.t).includes(q) || normalize(d.c).includes(q));
    }
    return list.slice(0, 8);
  }, [query, open]);

  return (
    <div style={{ position: 'relative' }} ref={wrapRef}>
      <div style={{ position: 'relative' }}>
        <Search size={15} style={{ position: 'absolute', left: 10, top: 12, color: 'var(--ink-faint)' }} />
        <input
          className="rd-input"
          style={{ paddingLeft: 32, paddingRight: value ? 30 : 12 }}
          placeholder={placeholder || 'Busca por número o palabra clave…'}
          value={query}
          onFocus={() => setOpen(true)}
          onChange={(e) => { setQuery(e.target.value); setOpen(true); if (!e.target.value) onChange(null); }}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
        />
        {value && (
          <button type="button" className="rd-picker-clear"
            onMouseDown={(e) => { e.preventDefault(); onChange(null); setQuery(''); }}>
            <X size={14} />
          </button>
        )}
      </div>
      {open && results.length > 0 && (
        <div className="rd-picker-list">
          {results.map((r) => (
            <button key={r.n} type="button" className="rd-picker-item"
              onMouseDown={(e) => { e.preventDefault(); onChange({ num: r.n, titulo: r.t }); setOpen(false); }}>
              <span className="rd-picker-num">{r.n}</span>
              <span style={{ minWidth: 0 }}>
                <div className="rd-picker-title">{r.t}</div>
                <div className="rd-picker-cat">{r.c}</div>
              </span>
            </button>
          ))}
        </div>
      )}
      {open && results.length === 0 && (
        <div className="rd-picker-list"><div className="rd-picker-empty">Sin resultados. Prueba con otra palabra o el número.</div></div>
      )}
    </div>
  );
}

const emptySalidaForm = () => ({ fecha: proximoSabado(), discursante: '', discurso: null, congregacion: '' });
const emptyVisitaForm = () => ({ fecha: proximoSabado(), nombre: '', congregacion: '', discurso: null });

export default function RegistroDiscursantesPaloSolo() {
  const [data, setData] = useState({ salidas: [], visitas: [] });
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState('salidas');
  const [infoOpen, setInfoOpen] = useState(false);
  const [confirmId, setConfirmId] = useState(null);
  const [salidaForm, setSalidaForm] = useState(emptySalidaForm());
  const [visitaForm, setVisitaForm] = useState(emptyVisitaForm());

  const load = useCallback(async (silent) => {
    try {
      if (!silent) setLoading(true);
      const stored = localStorage.getItem(STORAGE_KEY);
      const parsed = stored ? JSON.parse(stored) : { salidas: [], visitas: [] };
      setData({ salidas: parsed.salidas || [], visitas: parsed.visitas || [] });
    } catch (e) {
      // Si hay error, mantener datos previos
    } finally {
      setLoading(false);
    }
  }, []);

  const persist = useCallback(async (next) => {
    setSaving(true);
    setError(null);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      setData(next);
    } catch (e) {
      setError('No se pudo guardar. Revisa tu conexión e inténtalo de nuevo.');
    } finally {
      setSaving(false);
    }
  }, []);

  useEffect(() => {
    load(false);
  }, [load]);

  const speakerInfo = useMemo(() => ROSTER.find((r) => r.nombre === salidaForm.discursante), [salidaForm.discursante]);
  const speakerPreparados = useMemo(
    () => (speakerInfo ? speakerInfo.preparados.map((n) => DISCURSOS_POR_NUM[n]).filter(Boolean) : []),
    [speakerInfo]
  );

  const handleRegistrarSalida = async () => {
    if (!salidaForm.discursante || !salidaForm.fecha || saving) return;
    const nuevo = {
      id: nuevoId(),
      fecha: salidaForm.fecha,
      turno: speakerInfo ? speakerInfo.turno : null,
      discursante: salidaForm.discursante,
      discursoNum: salidaForm.discurso ? salidaForm.discurso.num : '',
      discursoTitulo: salidaForm.discurso ? salidaForm.discurso.titulo : '',
      congregacion: salidaForm.congregacion.trim(),
      registradoTs: Date.now(),
    };
    await persist({ ...data, salidas: [nuevo, ...data.salidas] });
    setSalidaForm(emptySalidaForm());
  };

  const handleRegistrarVisita = async () => {
    if (!visitaForm.nombre.trim() || !visitaForm.fecha || saving) return;
    const nuevo = {
      id: nuevoId(),
      fecha: visitaForm.fecha,
      nombre: visitaForm.nombre.trim(),
      congregacion: visitaForm.congregacion.trim(),
      discursoNum: visitaForm.discurso ? visitaForm.discurso.num : '',
      discursoTitulo: visitaForm.discurso ? visitaForm.discurso.titulo : '',
      registradoTs: Date.now(),
    };
    await persist({ ...data, visitas: [nuevo, ...data.visitas] });
    setVisitaForm(emptyVisitaForm());
  };

  const handleDelete = async (tipo, id) => {
    setConfirmId(null);
    if (tipo === 'salidas') await persist({ ...data, salidas: data.salidas.filter((e) => e.id !== id) });
    else await persist({ ...data, visitas: data.visitas.filter((e) => e.id !== id) });
  };

  const handleDescargar = () => {
    const wb = XLSX.utils.book_new();

    const salidasOrdenadas = [...data.salidas].sort((a, b) => a.fecha.localeCompare(b.fecha));
    const headerSalidas = ['Fecha', 'Discurso #', 'Titulo discurso', 'Congregación', ...ROSTER.map((r) => r.nombre)];
    const filasSalidas = salidasOrdenadas.map((e) => {
      const row = [fechaLegible(e.fecha), e.discursoNum || '', e.discursoTitulo || '', e.congregacion || ''];
      ROSTER.forEach((r) => row.push(r.nombre === e.discursante ? 'x' : ''));
      return row;
    });
    const wsSalidas = XLSX.utils.aoa_to_sheet([headerSalidas, ...filasSalidas]);
    wsSalidas['!cols'] = [{ wch: 11 }, { wch: 10 }, { wch: 34 }, { wch: 16 }, ...ROSTER.map(() => ({ wch: 14 }))];
    XLSX.utils.book_append_sheet(wb, wsSalidas, 'Salidas');

    const visitasOrdenadas = [...data.visitas].sort((a, b) => a.fecha.localeCompare(b.fecha));
    const headerVisitas = ['Fecha', '#', 'Número y Titulo discurso', 'Congregación', 'Discursante'];
    const filasVisitas = visitasOrdenadas.map((e) => [fechaLegible(e.fecha), e.discursoNum || '', e.discursoTitulo || '', e.congregacion || '', e.nombre || '']);
    const wsVisitas = XLSX.utils.aoa_to_sheet([headerVisitas, ...filasVisitas]);
    wsVisitas['!cols'] = [{ wch: 11 }, { wch: 6 }, { wch: 40 }, { wch: 16 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, wsVisitas, 'Nos visitan');

    const wsRoster = XLSX.utils.aoa_to_sheet([['Turno', 'Nombre', 'Rol', 'Teléfono'], ...ROSTER.map((r) => [r.turno, r.nombre, r.rol, r.tel])]);
    wsRoster['!cols'] = [{ wch: 7 }, { wch: 22 }, { wch: 22 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsRoster, 'Discursantes');

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `registro-discursantes-palo-solo-${fecha}.xlsx`);
  };

  const salidasList = [...data.salidas].sort((a, b) => b.fecha.localeCompare(a.fecha) || b.registradoTs - a.registradoTs);
  const visitasList = [...data.visitas].sort((a, b) => b.fecha.localeCompare(a.fecha) || b.registradoTs - a.registradoTs);

  return (
    <div className="rd-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .rd-root{
          --bg:#EEF3EF;--surface:#FFFFFF;--surface-2:#F6F9F7;--ink:#1F2A24;--ink-soft:#5C6B62;
          --ink-faint:#93A099;--amber:#D98E2B;--amber-dark:#B9721A;--amber-tint:#FBEDD8;
          --sage:#4C7A5E;--sage-dark:#3A5F49;--sage-tint:#E3EEE6;--line:#DCE3DE;--danger:#B3452D;--danger-tint:#F7E3DD;
          font-family:'Inter',sans-serif;background:var(--bg);color:var(--ink);min-height:100vh;padding:18px 12px 56px;
        }
        .rd-container{max-width:720px;margin:0 auto;}
        .rd-header{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:6px;flex-wrap:wrap;}
        .rd-title{font-family:'Fraunces',serif;font-weight:700;font-size:24px;letter-spacing:-0.01em;margin:0;}
        .rd-sub{font-size:12.5px;color:var(--ink-soft);margin-top:2px;}
        .rd-info-toggle{font-family:'Inter',sans-serif;font-size:12.5px;font-weight:600;color:var(--sage-dark);background:var(--sage-tint);border:none;border-radius:8px;padding:8px 10px;cursor:pointer;display:flex;align-items:center;gap:5px;flex-shrink:0;}
        .rd-info-panel{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:12px 14px;margin:10px 0 16px;font-size:13px;color:var(--ink-soft);display:flex;flex-direction:column;gap:6px;}
        .rd-info-row{display:flex;align-items:flex-start;gap:8px;}
        .rd-info-row a{color:var(--sage-dark);font-weight:600;text-decoration:none;}
        .rd-card{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:16px;margin-bottom:16px;}
        .rd-field-row{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px;}
        .rd-field{flex:1 1 160px;display:flex;flex-direction:column;gap:4px;min-width:0;}
        .rd-label{font-size:12px;font-weight:600;color:var(--ink-soft);}
        .rd-input,.rd-select{font-family:'Inter',sans-serif;font-size:15px;padding:10px 12px;border:1px solid var(--line);border-radius:8px;background:var(--surface-2);color:var(--ink);outline:none;width:100%;box-sizing:border-box;}
        .rd-input:focus-visible,.rd-select:focus-visible{border-color:var(--amber);background:var(--surface);box-shadow:0 0 0 3px var(--amber-tint);}
        .rd-turno-badge{display:inline-flex;align-items:center;gap:5px;font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:600;color:var(--sage-dark);background:var(--sage-tint);border-radius:6px;padding:2px 7px;margin-top:2px;}
        .rd-btn-primary{width:100%;font-family:'Inter',sans-serif;font-weight:700;font-size:15px;padding:13px 16px;background:var(--amber);color:#fff;border:none;border-radius:10px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;}
        .rd-btn-primary:hover:not(:disabled){background:var(--amber-dark);}
        .rd-btn-primary:disabled{opacity:.5;cursor:not-allowed;}
        .rd-btn-primary.sage{background:var(--sage);}
        .rd-btn-primary.sage:hover:not(:disabled){background:var(--sage-dark);}
        .rd-btn-primary:focus-visible{outline:3px solid var(--amber-dark);outline-offset:2px;}
        .rd-tabs{display:flex;gap:8px;margin-bottom:12px;}
        .rd-tab{flex:1;font-family:'Inter',sans-serif;font-weight:600;font-size:13.5px;padding:10px 6px;border-radius:10px;border:1px solid var(--line);background:var(--surface);color:var(--ink-soft);cursor:pointer;}
        .rd-tab.active{background:var(--sage);color:#fff;border-color:var(--sage);}
        .rd-tab:focus-visible{outline:3px solid var(--sage-dark);outline-offset:2px;}
        .rd-chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;}
        .rd-chip{font-family:'Inter',sans-serif;font-size:12px;font-weight:600;color:var(--sage-dark);background:var(--sage-tint);border:1px solid transparent;border-radius:999px;padding:6px 10px;cursor:pointer;text-align:left;max-width:100%;}
        .rd-chip:hover{border-color:var(--sage);}
        .rd-chip-hint{font-size:11.5px;color:var(--ink-faint);margin-bottom:4px;}
        .rd-picker-list{position:absolute;left:0;right:0;top:calc(100% + 4px);background:var(--surface);border:1px solid var(--line);border-radius:10px;box-shadow:0 8px 20px rgba(31,42,36,0.12);z-index:20;max-height:260px;overflow-y:auto;}
        .rd-picker-item{display:flex;align-items:flex-start;gap:8px;width:100%;text-align:left;padding:9px 10px;background:transparent;border:none;border-bottom:1px solid var(--line);cursor:pointer;}
        .rd-picker-item:last-child{border-bottom:none;}
        .rd-picker-item:hover{background:var(--surface-2);}
        .rd-picker-num{font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;color:var(--amber-dark);background:var(--amber-tint);border-radius:5px;padding:2px 6px;flex-shrink:0;margin-top:1px;}
        .rd-picker-title{font-size:13.5px;font-weight:600;color:var(--ink);line-height:1.3;}
        .rd-picker-cat{font-size:11px;color:var(--ink-faint);margin-top:1px;}
        .rd-picker-empty{padding:12px;font-size:13px;color:var(--ink-faint);text-align:center;}
        .rd-picker-clear{position:absolute;right:6px;top:6px;background:transparent;border:none;color:var(--ink-faint);cursor:pointer;padding:6px;border-radius:6px;}
        .rd-picker-clear:hover{color:var(--danger);}
        .rd-row{display:flex;align-items:center;gap:10px;padding:12px 14px;background:var(--surface);border:1px solid var(--line);border-radius:12px;margin-bottom:8px;flex-wrap:wrap;}
        .rd-row-main{flex:1 1 180px;min-width:0;}
        .rd-row-name{font-weight:700;font-size:14.5px;}
        .rd-row-sub{font-size:12.5px;color:var(--ink-soft);margin-top:1px;}
        .rd-stamp{font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:600;color:var(--ink-soft);border:1px dashed var(--ink-faint);border-radius:6px;padding:3px 7px;white-space:nowrap;}
        .rd-btn-icon{background:transparent;border:none;color:var(--ink-faint);cursor:pointer;padding:6px;border-radius:6px;display:flex;}
        .rd-btn-icon:hover{color:var(--danger);background:var(--danger-tint);}
        .rd-confirm{display:flex;align-items:center;gap:6px;}
        .rd-confirm-text{font-size:12px;color:var(--danger);font-weight:600;}
        .rd-confirm-btn{font-size:12px;font-weight:700;border:none;border-radius:6px;padding:5px 9px;cursor:pointer;}
        .rd-empty{text-align:center;padding:28px 16px;color:var(--ink-soft);font-size:13.5px;background:var(--surface-2);border:1px dashed var(--line);border-radius:12px;}
        .rd-footer{margin-top:18px;display:flex;justify-content:center;}
        .rd-btn-download{font-family:'Inter',sans-serif;font-weight:700;font-size:14px;padding:12px 22px;background:var(--ink);color:#fff;border:none;border-radius:10px;cursor:pointer;display:flex;align-items:center;gap:8px;}
        .rd-btn-download:hover{background:#0f1712;}
        .rd-error{display:flex;align-items:center;gap:8px;background:var(--danger-tint);color:var(--danger);border-radius:10px;padding:10px 14px;font-size:13px;margin-bottom:14px;font-weight:600;}
        .rd-saving{font-size:12px;color:var(--ink-faint);text-align:center;margin-top:8px;}
      `}</style>

      <div className="rd-container">
        <div className="rd-header">
          <div>
            <h1 className="rd-title">Registro de Discursantes</h1>
            <div className="rd-sub">Congregación Palo Solo</div>
          </div>
          <button className="rd-info-toggle" onClick={() => setInfoOpen((v) => !v)}>
            <ChevronDown size={14} style={{ transform: infoOpen ? 'rotate(180deg)' : 'none', transition: 'transform .15s' }} />
            Info del salón
          </button>
        </div>

        {infoOpen && (
          <div className="rd-info-panel">
            <div className="rd-info-row"><MapPin size={15} style={{ marginTop: 1 }} />
              <span>C. Benito Juárez #22, Colonia Palo Solo, 52778, Huixquilucan de Degollado, Méx.
                {' '}<a href="https://maps.app.goo.gl/wx5GSCFr41RNDRKL7" target="_blank" rel="noreferrer">Ver en Google Maps</a>
              </span>
            </div>
            <div className="rd-info-row"><Clock3 size={15} /> Sábados 07:00 pm</div>
            <div className="rd-info-row"><Phone size={15} /> Coordinador de discursos: Javier Santos · 55 8469 2565</div>
          </div>
        )}

        {error && <div className="rd-error"><AlertCircle size={16} />{error}</div>}

        <div className="rd-tabs">
          <button className={`rd-tab ${tab === 'salidas' ? 'active' : ''}`} onClick={() => setTab('salidas')}>Discursantes que salen ({data.salidas.length})</button>
          <button className={`rd-tab ${tab === 'visitas' ? 'active' : ''}`} onClick={() => setTab('visitas')}>Discursantes que nos visitan ({data.visitas.length})</button>
        </div>

        {tab === 'salidas' ? (
          <>
            <div className="rd-card">
              <div className="rd-field-row">
                <div className="rd-field">
                  <label className="rd-label" htmlFor="s-fecha">Fecha</label>
                  <input id="s-fecha" type="date" className="rd-input" value={salidaForm.fecha}
                    onChange={(e) => setSalidaForm({ ...salidaForm, fecha: e.target.value })} />
                </div>
                <div className="rd-field">
                  <label className="rd-label" htmlFor="s-disc">Discursante</label>
                  <select id="s-disc" className="rd-select" value={salidaForm.discursante}
                    onChange={(e) => setSalidaForm({ ...salidaForm, discursante: e.target.value, discurso: null })}>
                    <option value="">Selecciona…</option>
                    {ROSTER.map((r) => <option key={r.nombre} value={r.nombre}>{r.turno}. {r.nombre}</option>)}
                  </select>
                  {speakerInfo && <span className="rd-turno-badge">Turno {speakerInfo.turno} · {speakerInfo.rol}</span>}
                </div>
              </div>

              {salidaForm.discursante && (
                <div className="rd-field-row" style={{ flexDirection: 'column' }}>
                  {speakerPreparados.length > 0 && (
                    <div style={{ marginBottom: 4 }}>
                      <div className="rd-chip-hint">Discursos que ya tiene preparados:</div>
                      <div className="rd-chips">
                        {speakerPreparados.map((d) => (
                          <button key={d.n} type="button" className="rd-chip"
                            onClick={() => setSalidaForm({ ...salidaForm, discurso: { num: d.n, titulo: d.t } })}>
                            {d.n} — {d.t.length > 38 ? d.t.slice(0, 38) + '…' : d.t}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                  <div className="rd-field" style={{ flexBasis: '100%' }}>
                    <label className="rd-label">O busca cualquier otro discurso del catálogo</label>
                    <TalkPicker value={salidaForm.discurso} onChange={(d) => setSalidaForm({ ...salidaForm, discurso: d })} />
                  </div>
                </div>
              )}

              <div className="rd-field-row">
                <div className="rd-field" style={{ flexBasis: '100%' }}>
                  <label className="rd-label" htmlFor="s-cong">Congregación destino</label>
                  <input id="s-cong" className="rd-input" list="congs" placeholder="Ej. Parque la punta"
                    value={salidaForm.congregacion} onChange={(e) => setSalidaForm({ ...salidaForm, congregacion: e.target.value })} />
                  <datalist id="congs">{CONGREGACIONES_SUGERIDAS.map((c) => <option key={c} value={c} />)}</datalist>
                </div>
              </div>

              <button className="rd-btn-primary" onClick={handleRegistrarSalida} disabled={!salidaForm.discursante || saving}>
                <LogOut size={18} /> Registrar salida
              </button>
              {saving && <div className="rd-saving">Guardando…</div>}
            </div>

            {loading ? <div className="rd-empty">Cargando registro…</div> :
              salidasList.length === 0 ? <div className="rd-empty">Aún no hay salidas registradas. Llena el formulario de arriba para comenzar.</div> :
              salidasList.map((e) => (
                <div className="rd-row" key={e.id}>
                  <div className="rd-row-main">
                    <div className="rd-row-name">{e.discursante} {e.turno ? <span className="rd-turno-badge">Turno {e.turno}</span> : null}</div>
                    <div className="rd-row-sub">{e.discursoNum ? `#${e.discursoNum} — ` : ''}{e.discursoTitulo || 'Sin discurso especificado'}{e.congregacion ? ` · a ${e.congregacion}` : ''}</div>
                  </div>
                  <span className="rd-stamp">{fechaLegible(e.fecha)}</span>
                  {confirmId === e.id ? (
                    <div className="rd-confirm">
                      <span className="rd-confirm-text">¿Eliminar?</span>
                      <button className="rd-confirm-btn" style={{ background: 'var(--danger)', color: '#fff' }} onClick={() => handleDelete('salidas', e.id)}>Sí</button>
                      <button className="rd-confirm-btn" style={{ background: 'var(--surface-2)', color: 'var(--ink-soft)' }} onClick={() => setConfirmId(null)}>No</button>
                    </div>
                  ) : <button className="rd-btn-icon" onClick={() => setConfirmId(e.id)} title="Eliminar"><Trash2 size={16} /></button>}
                </div>
              ))
            }
          </>
        ) : (
          <>
            <div className="rd-card">
              <div className="rd-field-row">
                <div className="rd-field">
                  <label className="rd-label" htmlFor="v-fecha">Fecha</label>
                  <input id="v-fecha" type="date" className="rd-input" value={visitaForm.fecha}
                    onChange={(e) => setVisitaForm({ ...visitaForm, fecha: e.target.value })} />
                </div>
                <div className="rd-field">
                  <label className="rd-label" htmlFor="v-nombre">Nombre del discursante visitante</label>
                  <input id="v-nombre" className="rd-input" placeholder="Escribe el nombre" value={visitaForm.nombre}
                    onChange={(e) => setVisitaForm({ ...visitaForm, nombre: e.target.value })} />
                </div>
              </div>
              <div className="rd-field-row">
                <div className="rd-field" style={{ flexBasis: '100%' }}>
                  <label className="rd-label">Discurso</label>
                  <TalkPicker value={visitaForm.discurso} onChange={(d) => setVisitaForm({ ...visitaForm, discurso: d })} />
                </div>
              </div>
              <div className="rd-field-row">
                <div className="rd-field" style={{ flexBasis: '100%' }}>
                  <label className="rd-label" htmlFor="v-cong">Congregación de origen</label>
                  <input id="v-cong" className="rd-input" list="congs2" placeholder="Ej. Parque la punta"
                    value={visitaForm.congregacion} onChange={(e) => setVisitaForm({ ...visitaForm, congregacion: e.target.value })} />
                  <datalist id="congs2">{CONGREGACIONES_SUGERIDAS.map((c) => <option key={c} value={c} />)}</datalist>
                </div>
              </div>
              <button className="rd-btn-primary sage" onClick={handleRegistrarVisita} disabled={!visitaForm.nombre.trim() || saving}>
                <LogIn size={18} /> Registrar visita
              </button>
              {saving && <div className="rd-saving">Guardando…</div>}
            </div>

            {loading ? <div className="rd-empty">Cargando registro…</div> :
              visitasList.length === 0 ? <div className="rd-empty">Aún no hay visitas registradas. Llena el formulario de arriba para comenzar.</div> :
              visitasList.map((e) => (
                <div className="rd-row" key={e.id}>
                  <div className="rd-row-main">
                    <div className="rd-row-name">{e.nombre}</div>
                    <div className="rd-row-sub">{e.discursoNum ? `#${e.discursoNum} — ` : ''}{e.discursoTitulo}{e.congregacion ? ` · de ${e.congregacion}` : ''}</div>
                  </div>
                  <span className="rd-stamp">{fechaLegible(e.fecha)}</span>
                  {confirmId === e.id ? (
                    <div className="rd-confirm">
                      <span className="rd-confirm-text">¿Eliminar?</span>
                      <button className="rd-confirm-btn" style={{ background: 'var(--danger)', color: '#fff' }} onClick={() => handleDelete('visitas', e.id)}>Sí</button>
                      <button className="rd-confirm-btn" style={{ background: 'var(--surface-2)', color: 'var(--ink-soft)' }} onClick={() => setConfirmId(null)}>No</button>
                    </div>
                  ) : <button className="rd-btn-icon" onClick={() => setConfirmId(e.id)} title="Eliminar"><Trash2 size={16} /></button>}
                </div>
              ))
            }
          </>
        )}

        <div className="rd-footer">
          <button className="rd-btn-download" onClick={handleDescargar}><Download size={17} /> Descargar Excel</button>
        </div>
      </div>
    </div>
  );
}