// import { ITimeService } from "./ITimeService";

// export class TimeService implements ITimeService {
//     private o = {
//         second: 1000,
//         minute: 60 * 1000,
//         hour: 60 * 1000 * 60,
//         day: 24 * 60 * 1000 * 60,
//         week: 7 * 24 * 60 * 1000 * 60,
//         month: 30 * 24 * 60 * 1000 * 60,
//         year: 365 * 24 * 60 * 1000 * 60
//     };
//     public timefriendly(s: string) {
//         let t = s.match(/(\d).([a-z]*?)s?$/);
//         return (t[1] * eval(this.o[t[2]])).toString();
//     }
//     public mintoread(text: string, altcmt, wpm) {
//         let m = Math.round(text.split(' ').length / (wpm || 200));
//         return (m || '< 1') + (altcmt || ' min to read');
//     }
//     public today() {
//         let now = new Date();
//         let Weekday = new Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
//         let Month = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
//         return Weekday[now.getDay()] + ", " + Month[now.getMonth()] + " " + now.getDate() + ", " + now.getFullYear();
//     }
//     public pl(v, n) {
//         return (s === undefined) ? n + ' ' + v + (n > 1 ? 's' : '') + dir : n + v.substring(0, 1)
//     }
//     public ago(nd, s) {
//         let r = Math.round,
//             dir = ' ago',

//             ts = Date.now() - new Date(nd).getTime(),
//             ii;
//         if (ts < 0) {
//             ts *= -1;
//             dir = ' from now';
//         }
//         for (let i in this.o) {
//             if (r(ts) < this.o[i]) return pl(ii || 'm', r(ts / (this.o[ii] || 1)))
//             ii = i;
//         }
//         return pl(i, r(ts / this.o[i]));
//     }
// }