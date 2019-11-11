import * as React from 'react';
import styles from './NewsRotator.module.scss';
import { INewsRotatorProps } from './INewsRotatorProps';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import "../Styles/NewsRotator.scss";

export default class NewsRotator extends React.Component<INewsRotatorProps, {}> {
  
  private items:any;
  private itemsArr = [];
  
  public render(): React.ReactElement<INewsRotatorProps> {
    //console.info(this.props);
    let self = this;

    this.props.listItems.forEach((item,i) => {
      if(this.props.listItems[i].BannerImageUrl.Url.indexOf('Resolution') == -1){
        this.props.listItems[i].BannerImageUrl.Url += this.props.listItems[i].BannerImageUrl.Url.indexOf('?') > -1 ? '&Resolution=3' : '?Resolution=3';
      }
    });

    this.itemsArr = [];
    this.items = this.props.listItems.map((item, key) =>
      { if(item) {
        this.itemsArr.push(key);
        return <div className={styles.SlickSlideItem}><a href={item.FileRef} title={item.Title} target="_blank"><div style={{backgroundImage: `url(${item.BannerImageUrl.Url})`}}></div></a></div>;
      }}
    );

    function getButtonClass(){
      var buttonwidths = [styles.SlickDots1,styles.SlickDots2,styles.SlickDots3,styles.SlickDots4,styles.SlickDots5];
      return buttonwidths[self.itemsArr.length-1];
    }

    const settings = {
      dots: true,
      infinite: true,
      speed: this.props.animationSpeed,
      arrows: false,
      slidesToShow: 1,
      slidesToScroll: 1,
      fade: this.props.fade,
      autoplay: this.props.autoplay,
      autoplaySpeed: this.props.autoplaySpeed,
      dotsClass: getButtonClass(), 
      appendDots: dots => (
          <ul> {dots} </ul>
      ),
      customPaging: i => (
          <div className={styles.SlickDotsListItem} title={this.props.listItems[this.itemsArr[i]].Title}> {this.props.listItems[this.itemsArr[i]].Title} </div>
      )
    };

    return (
      <div className={styles.newsRotator}>
        {this.props.title != '' ?
          <h2 className={styles.WPtitle}><span>{this.props.title}</span></h2>
        : ''}
        <Slider {...settings} className={styles.SlickSlider}>
          {this.items}
        </Slider>
      </div>
    );
  }
}
