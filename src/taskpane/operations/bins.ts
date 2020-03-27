import { hsl } from "d3";
/* global console */

export default class Bins {
  private minDomain: number;
  private maxDomain: number;
  private width: number;
  private nrOfBins: number;
  public bins: Bin[]

  constructor(minDomain: number, maxDomain: number, width: number) {
    this.minDomain = minDomain;
    this.maxDomain = maxDomain;
    this.width = width;
    this.nrOfBins = (this.maxDomain - this.minDomain) / this.width;
    this.bins = new Array<Bin>();
  }

  createBins(data: number[]) {

    let i = 0;
    this.bins = new Array<Bin>();

    while (i < this.nrOfBins) {

      let bin: Bin = { x0: 0, x1: 0, length: 0, samples: new Array<number>() };

      bin.x0 = this.minDomain + this.width * i;
      bin.x1 = bin.x0 + this.width;
      bin.length = 0;
      bin.samples = new Array<number>();
      this.bins.push(bin);
      i++;
    }

    data.forEach((element: number) => {
      this.bins.forEach((bin: Bin) => {
        if (element >= bin.x0 && element < bin.x1) {
          bin.samples.push(element);
          bin.length++;
        }
      });
    })
    return this.bins;
  }

  getTickValues() {
    let ticks = new Array<number>();
    let i = this.minDomain;

    while (i <= this.maxDomain) {
      ticks.push(i);
      i = i + 6;
    }

    return ticks;
  }

  generateBlueColors() {

    let blueColors = [];
    let i = 0;
    while (i <= this.nrOfBins) {
      let color = hsl(198, 1, 0.97 * (1 - i / 15));
      blueColors.push(color.hex());
      i++;
    }

    return blueColors;
  }

  generateOrangeColors() {

    let orangeColors = [];
    let i = 0;
    while (i <= this.nrOfBins) {
      let color = hsl(34, 1, 0.97 * (1 - i / 15));
      orangeColors.push(color.hex());
      i++;
    }

    return orangeColors;
  }
}

interface Bin {
  x0: number;
  x1: number;
  length: number;
  samples: number[];
}


